import type {
	IDataObject,
	IHttpRequestOptions,
	IN8nHttpFullResponse,
	INodeExecutionData,
	INodeProperties,
	INodePropertyOptions,
} from 'n8n-workflow';
import { beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('../GenericFunctions', async (importOriginal) => {
	const actual = await importOriginal<typeof import('../GenericFunctions')>();
	return { ...actual, microsoftApiRequest: vi.fn(async () => ({})) };
});

const { microsoftApiRequest } = await import('../GenericFunctions');
const { groupFields, groupOperations, licenseFields, licenseOperations, userFields, userOperations } =
	await import('./index');

type Handler = (...args: unknown[]) => Promise<unknown>;
interface Collected {
	owner: string;
	fn: Handler;
}

/** Walks a property tree, including nested collection/fixedCollection options. */
function collect(properties: INodeProperties[], kind: 'preSend' | 'postReceive'): Collected[] {
	const found: Collected[] = [];

	const visit = (property: INodeProperties | INodePropertyOptions, path: string) => {
		const owner = path ? `${path}.${property.name}` : property.name;
		const routing = (property as INodeProperties).routing;

		const handlers =
			kind === 'preSend' ? routing?.send?.preSend : (routing?.output?.postReceive as unknown[]);

		for (const handler of handlers ?? []) {
			if (typeof handler === 'function') {
				found.push({ owner, fn: handler as Handler });
			}
		}

		const nested = (property as INodeProperties).options as
			| Array<INodeProperties | INodePropertyOptions>
			| undefined;
		for (const child of nested ?? []) {
			if (typeof child === 'object' && 'name' in child) visit(child, owner);
		}

		for (const value of (property as INodeProperties).values ?? []) {
			visit(value, owner);
		}
	};

	for (const property of properties) visit(property, '');
	return found;
}

const ALL_PROPERTIES = [
	...groupOperations,
	...groupFields,
	...licenseOperations,
	...licenseFields,
	...userOperations,
	...userFields,
];

const PRE_SENDS = collect(ALL_PROPERTIES, 'preSend');
const POST_RECEIVES = collect(ALL_PROPERTIES, 'postReceive');
const GROUP_POST_RECEIVES = collect(groupFields, 'postReceive');
const USER_POST_RECEIVES = collect(userFields, 'postReceive');

function handlerFor(collected: Collected[], owner: string): Handler {
	const found = collected.find((h) => h.owner.includes(owner));
	if (!found) throw new Error(`no handler owned by ${owner}`);
	return found.fn;
}

function context(parameters: Record<string, unknown>) {
	return {
		getNode: () => ({ name: 'Microsoft 365 Admin', type: 'microsoft365Admin' }),
		getNodeParameter: (name: string) => parameters[name] ?? '',
		helpers: { httpRequestWithAuthentication: vi.fn() },
		getCredentials: vi.fn(async () => ({})),
	};
}

function requestOptions(body: IDataObject = {}): IHttpRequestOptions {
	return { url: 'https://graph.microsoft.com/v1.0/groups', body } as IHttpRequestOptions;
}

const RESPONSE = { statusCode: 200, body: {}, headers: {} } as unknown as IN8nHttpFullResponse;

beforeEach(() => {
	vi.mocked(microsoftApiRequest).mockClear();
	vi.mocked(microsoftApiRequest).mockImplementation(async () => ({}));
});

describe('inline routing handlers', () => {
	it('finds the handlers embedded in the descriptions', () => {
		expect(PRE_SENDS.length).toBeGreaterThanOrEqual(14);
		expect(POST_RECEIVES.length).toBeGreaterThanOrEqual(2);
	});

	// Values that satisfy every validator, so the happy path of each handler runs.
	const BENIGN = {
		displayName: 'Platform Team',
		mailNickname: 'platform',
		userPrincipalName: 'platform.team',
		'updateFields.displayName': 'Platform Team',
		'updateFields.mailNickname': 'platform',
		'updateFields.userPrincipalName': 'platform.team',
	};

	it('every preSend returns request options for benign input', async () => {
		for (const { owner, fn } of PRE_SENDS) {
			const result = await fn.call(context(BENIGN), requestOptions());
			expect(result, owner).toBeDefined();
		}
	});

	it('every preSend rejects an over-long value rather than letting Graph 400', async () => {
		const tooLong = 'x'.repeat(300);
		let rejections = 0;

		for (const { fn } of PRE_SENDS) {
			const ctx = context({
				displayName: tooLong,
				mailNickname: tooLong,
				'updateFields.displayName': tooLong,
				'updateFields.mailNickname': tooLong,
				description: tooLong,
			});

			try {
				await fn.call(ctx, requestOptions());
			} catch {
				rejections++;
			}
		}

		expect(rejections).toBeGreaterThan(0);
	});

	it('every postReceive returns the items it was given', async () => {
		for (const { owner, fn } of POST_RECEIVES) {
			const items: INodeExecutionData[] = [{ json: { id: 'g1' }, index: 0 }];
			const result = (await fn.call(context({}), items, RESPONSE)) as INodeExecutionData[];
			expect(result, owner).toHaveLength(1);
		}
	});
});

describe('group create validation', () => {
	function preSendFor(field: string): Handler {
		const found = PRE_SENDS.find((h) => h.owner === field || h.owner.endsWith(`.${field}`));
		if (!found) throw new Error(`no preSend for ${field}`);
		return found.fn;
	}

	it('marks a security group as mail-disabled and security-enabled', async () => {
		const options = requestOptions();
		await preSendFor('groupType').call(context({ groupType: '' }), options);

		expect(options.body).toMatchObject({ mailEnabled: false, securityEnabled: true });
	});

	it('adds the Unified group type for a Microsoft 365 group', async () => {
		const options = requestOptions();
		await preSendFor('groupType').call(context({ groupType: 'Unified' }), options);

		expect((options.body as IDataObject).groupTypes).toEqual(['Unified']);
	});

	it('rejects a display name longer than 256 characters', async () => {
		await expect(
			preSendFor('displayName').call(
				context({ displayName: 'x'.repeat(257) }),
				requestOptions(),
			),
		).rejects.toThrow(/maximum length of 256/);
	});

	it('accepts a display name at the 256-character limit', async () => {
		await expect(
			preSendFor('displayName').call(
				context({ displayName: 'x'.repeat(256) }),
				requestOptions(),
			),
		).resolves.toBeDefined();
	});

	it('rejects a mail nickname longer than 64 characters', async () => {
		await expect(
			preSendFor('mailNickname').call(
				context({ mailNickname: 'x'.repeat(65) }),
				requestOptions(),
			),
		).rejects.toThrow(/maximum length of 64/);
	});

	it('tells the user to drop the domain when a full address is pasted', async () => {
		await expect(
			preSendFor('mailNickname').call(
				context({ mailNickname: 'platform@contoso.com' }),
				requestOptions(),
			),
		).rejects.toThrow(/only include the local-part/);
	});

	it.each(['bad nickname', 'bad,nickname', 'bad<nickname', 'bad;nickname'])(
		'rejects the reserved character in "%s"',
		async (mailNickname) => {
			await expect(
				preSendFor('mailNickname').call(context({ mailNickname }), requestOptions()),
			).rejects.toThrow(/ASCII character set/);
		},
	);

	it('rejects characters Entra disallows in a user principal name', async () => {
		const found = PRE_SENDS.find((h) => h.owner.endsWith('userPrincipalName'));
		expect(found).toBeDefined();

		await expect(
			found!.fn.call(context({ userPrincipalName: 'bad name' }), requestOptions()),
		).rejects.toThrow(/Only the following characters are allowed/);
	});

	it('accepts a well-formed mail nickname', async () => {
		await expect(
			preSendFor('mailNickname').call(context({ mailNickname: 'engineering' }), requestOptions()),
		).resolves.toBeDefined();
	});
});

describe('group create follow-up patch', () => {
	function createPostReceive(): Handler {
		return handlerFor(GROUP_POST_RECEIVES, 'additionalFields');
	}

	const items = (): INodeExecutionData[] => [{ json: { id: 'group-1' }, index: 0 }];

	it('skips the extra request when no additional fields were supplied', async () => {
		await createPostReceive().call(context({ additionalFields: {} }), items(), RESPONSE);
		expect(microsoftApiRequest).not.toHaveBeenCalled();
	});

	it('patches the created group with the remaining fields', async () => {
		const result = (await createPostReceive().call(
			context({ additionalFields: { description: 'Platform team' } }),
			items(),
			RESPONSE,
		)) as INodeExecutionData[];

		expect(microsoftApiRequest).toHaveBeenCalledWith(
			'PATCH',
			'/groups/group-1',
			expect.objectContaining({ description: 'Platform team' }),
		);
		expect(result[0].json).toMatchObject({ description: 'Platform team' });
	});

	it('does not resend fields that were already applied at creation', async () => {
		await createPostReceive().call(
			context({
				additionalFields: {
					isAssignableToRole: true,
					membershipRule: 'rule',
					membershipRuleProcessingState: 'On',
				},
			}),
			items(),
			RESPONSE,
		);

		expect(microsoftApiRequest).not.toHaveBeenCalled();
	});

	it('deletes the half-created group when the follow-up patch fails', async () => {
		vi.mocked(microsoftApiRequest).mockImplementation(async (method: unknown) => {
			if (method === 'PATCH') throw new Error('patch failed');
			return {};
		});

		await expect(
			createPostReceive().call(
				context({ additionalFields: { description: 'Platform team' } }),
				items(),
				RESPONSE,
			),
		).rejects.toThrow('patch failed');

		expect(microsoftApiRequest).toHaveBeenCalledWith('DELETE', '/groups/group-1');
	});

	it('still surfaces the original failure when the rollback delete also fails', async () => {
		vi.mocked(microsoftApiRequest).mockImplementation(async () => {
			throw new Error('patch failed');
		});

		await expect(
			createPostReceive().call(
				context({ additionalFields: { description: 'Platform team' } }),
				items(),
				RESPONSE,
			),
		).rejects.toThrow('patch failed');
	});
});

describe('group update follow-up patch', () => {
	const handler = () => handlerFor(GROUP_POST_RECEIVES, 'updateFields');
	const items = (): INodeExecutionData[] => [{ json: {}, index: 0 }];

	async function run(updateFields: IDataObject) {
		return await handler().call(
			context({ 'group.value': 'group-9', updateFields }),
			items(),
			RESPONSE,
		);
	}

	it('makes no extra request for ordinary properties', async () => {
		await run({ displayName: 'Renamed' });
		expect(microsoftApiRequest).not.toHaveBeenCalled();
	});

	it('sends Exchange-backed properties in their own request', async () => {
		await run({ displayName: 'Renamed', allowExternalSenders: true });

		expect(microsoftApiRequest).toHaveBeenCalledExactlyOnceWith(
			'PATCH',
			'/groups/group-9',
			{ allowExternalSenders: true },
		);
	});

	it('groups multiple separate-only properties into a single request', async () => {
		await run({ allowExternalSenders: true, autoSubscribeNewMembers: false });

		expect(microsoftApiRequest).toHaveBeenCalledExactlyOnceWith('PATCH', '/groups/group-9', {
			allowExternalSenders: true,
			autoSubscribeNewMembers: false,
		});
	});

	it('returns the items unchanged', async () => {
		const result = (await run({ allowExternalSenders: true })) as INodeExecutionData[];
		expect(result).toHaveLength(1);
	});
});

describe('user create follow-up patch', () => {
	const handler = () => handlerFor(USER_POST_RECEIVES, 'additionalFields');
	const items = (): INodeExecutionData[] => [{ json: { id: 'user-1' }, index: 0 }];

	/** Stands in for a luxon DateTime; only these two methods are used. */
	const dateTime = (iso: string) => ({ toUTC: () => ({ toISO: () => iso }) });

	async function run(additionalFields: IDataObject) {
		return (await handler().call(
			context({ additionalFields }),
			items(),
			RESPONSE,
		)) as INodeExecutionData[];
	}

	function patchBodies(): IDataObject[] {
		return vi
			.mocked(microsoftApiRequest)
			.mock.calls.filter((c) => c[0] === 'PATCH')
			.map((c) => c[2] as IDataObject);
	}

	it('makes no follow-up request when nothing extra was supplied', async () => {
		await run({});
		expect(microsoftApiRequest).not.toHaveBeenCalled();
	});

	it('converts date fields to UTC ISO strings', async () => {
		await run({
			birthday: dateTime('1990-01-01T00:00:00.000Z'),
			employeeHireDate: dateTime('2020-06-01T00:00:00.000Z'),
			employeeLeaveDateTime: dateTime('2026-01-01T00:00:00.000Z'),
		});

		const merged = Object.assign({}, ...patchBodies()) as IDataObject;
		expect(merged.employeeHireDate).toBe('2020-06-01T00:00:00.000Z');
		expect(merged.employeeLeaveDateTime).toBe('2026-01-01T00:00:00.000Z');
		expect(merged.birthday).toBe('1990-01-01T00:00:00.000Z');
	});

	it('wraps a single business phone into the array Graph expects', async () => {
		await run({ businessPhones: '+370 600 00000' });
		expect(patchBodies()[0].businessPhones).toEqual(['+370 600 00000']);
	});

	it('unwraps the employee org data collection', async () => {
		await run({ employeeOrgData: { employeeOrgValues: { division: 'Platform' } } });
		expect(patchBodies()[0].employeeOrgData).toEqual({ division: 'Platform' });
	});

	it('flattens password policies into a comma-separated string', async () => {
		await run({ passwordPolicies: ['DisableStrongPassword', 'DisablePasswordExpiration'] });
		expect(patchBodies()[0].passwordPolicies).toBe(
			'DisableStrongPassword,DisablePasswordExpiration',
		);
	});

	it('moves a force-change-password choice into the password profile', async () => {
		await run({ forceChangePassword: 'forceChangePasswordNextSignIn' });

		const [body] = patchBodies();
		expect(body.passwordProfile).toMatchObject({ forceChangePasswordNextSignIn: true });
		expect(body.forceChangePassword).toBeUndefined();
	});

	it('handles the MFA variant of force-change-password', async () => {
		await run({ forceChangePassword: 'forceChangePasswordNextSignInWithMfa' });

		const [body] = patchBodies();
		expect(body.passwordProfile).toMatchObject({ forceChangePasswordNextSignInWithMfa: true });
	});

	it('sends properties Graph only accepts alone in their own request', async () => {
		await run({ aboutMe: 'Engineer', skills: ['n8n'], department: 'Platform' });

		const bodies = patchBodies();
		expect(bodies).toHaveLength(2);

		const separate = bodies.find((b) => 'aboutMe' in b);
		const regular = bodies.find((b) => 'department' in b);

		expect(separate).toEqual({ aboutMe: 'Engineer', skills: ['n8n'] });
		expect(regular).toEqual({ department: 'Platform' });
	});

	it('merges everything it applied back onto the output item', async () => {
		const result = await run({ department: 'Platform', aboutMe: 'Engineer' });

		expect(result[0].json).toMatchObject({
			id: 'user-1',
			department: 'Platform',
			aboutMe: 'Engineer',
		});
	});

	it('deletes the half-created user when the follow-up patch fails', async () => {
		vi.mocked(microsoftApiRequest).mockImplementation(async (method: unknown) => {
			if (method === 'PATCH') throw new Error('patch failed');
			return {};
		});

		await expect(run({ department: 'Platform' })).rejects.toThrow('patch failed');
		expect(microsoftApiRequest).toHaveBeenCalledWith('DELETE', '/users/user-1');
	});
});
