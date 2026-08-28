import { describe, expect, it } from 'vitest';

import { operationContext } from '../operation.fixtures';
import { execute } from './resetPassword.operation';

const USER = '02bd9fd6-8f93-4758-87c3-1fb73740a315';

async function run(parameters: Record<string, unknown>) {
	const { ctx, requests } = operationContext({ parameters });
	const output = await execute.call(ctx, 0);
	return { output, requests };
}

describe('reset password', () => {
	it('patches passwordProfile with the password it was given', async () => {
		const { output, requests } = await run({
			user: USER,
			options: { password: 'Example-Horse-42' },
		});

		expect(requests[0]).toMatchObject({
			method: 'PATCH',
			url: `https://graph.microsoft.com/v1.0/users/${USER}`,
		});
		expect(requests[0].body).toEqual({
			passwordProfile: { password: 'Example-Horse-42', forceChangePasswordNextSignIn: true },
		});
		expect(output[0].json).toMatchObject({
			id: USER,
			passwordReset: true,
			password: 'Example-Horse-42',
			generated: false,
		});
	});

	it('generates a password when none is given and hands it back', async () => {
		const { output, requests } = await run({ user: USER });

		const body = requests[0].body as { passwordProfile: { password: string } };
		expect(output[0].json.generated).toBe(true);
		expect(output[0].json.password).toBe(body.passwordProfile.password);
		expect(String(output[0].json.password)).toHaveLength(16);
	});

	it('honours the requested length of a generated password', async () => {
		const { output } = await run({ user: USER, options: { passwordLength: 32 } });
		expect(String(output[0].json.password)).toHaveLength(32);
	});

	it('keeps a length Graph would reject inside the allowed range', async () => {
		const { output } = await run({ user: USER, options: { passwordLength: 2 } });
		expect(String(output[0].json.password).length).toBeGreaterThanOrEqual(8);
	});

	it('asks for MFA before the change when told to', async () => {
		const { requests } = await run({
			user: USER,
			options: { forceChangePassword: 'forceChangePasswordNextSignInWithMfa' },
		});

		expect(requests[0].body).toMatchObject({
			passwordProfile: { forceChangePasswordNextSignInWithMfa: true },
		});
	});

	it('leaves the password in place when no change is required', async () => {
		const { output, requests } = await run({
			user: USER,
			options: { forceChangePassword: 'never' },
		});

		expect(requests[0].body).toMatchObject({
			passwordProfile: { forceChangePasswordNextSignIn: false },
		});
		expect(output[0].json.forceChangePasswordNextSignIn).toBe(false);
	});

	it('refuses a user value that would reshape the request path', async () => {
		const { ctx, httpRequestWithAuthentication } = operationContext({
			parameters: { user: 'alice/authentication/methods?x=' },
		});

		await expect(execute.call(ctx, 0)).rejects.toThrow(/characters that are not allowed/);
		expect(httpRequestWithAuthentication).not.toHaveBeenCalled();
	});

	it('rejects an item with no user', async () => {
		const { ctx } = operationContext({ parameters: { user: '' } });
		await expect(execute.call(ctx, 0)).rejects.toThrow(/No user was given/);
	});

	it('pairs its output back to the item it ran for', async () => {
		const { ctx } = operationContext({ parameters: { user: USER } });
		const output = await execute.call(ctx, 2);
		expect(output[0].pairedItem).toEqual({ item: 2 });
	});
});
