import { NodeOperationError, type INode } from 'n8n-workflow';

/**
 * Characters Entra accepts in a group's mail alias: printable ASCII, minus the punctuation
 * an address cannot carry. Space is in neither set, so it is rejected by the range itself.
 *
 * The range starts at 0x21 rather than 0x00 — a control character is never a legitimate
 * mail alias, and Graph rejects one anyway, so allowing them only meant this check passed
 * a value the API would refuse.
 */
const ASCII_MAIL_NICKNAME = /^((?![@()[\]"\\;:<>,])[\x21-\x7E])*$/;

const MAX_DISPLAY_NAME = 256;
const MAX_MAIL_NICKNAME = 64;

/**
 * Checks the two group fields Graph would otherwise reject with a message that names
 * neither the parameter nor the limit it broke. Shared by Create and Update, which take
 * the same values under different parameter names.
 */
export function validateGroupNames(
	node: INode,
	{ displayName = '', mailNickname = '' }: { displayName?: string; mailNickname?: string },
	itemIndex: number,
): void {
	if (displayName.length > MAX_DISPLAY_NAME) {
		throw new NodeOperationError(
			node,
			`'Display Name' should have a maximum length of ${MAX_DISPLAY_NAME}`,
			{ itemIndex },
		);
	}

	if (mailNickname.includes('@')) {
		throw new NodeOperationError(
			node,
			`'Group Email Address' should only include the local-part of the email address, without ${mailNickname.slice(mailNickname.indexOf('@'))}`,
			{ itemIndex },
		);
	}

	if (mailNickname.length > MAX_MAIL_NICKNAME) {
		throw new NodeOperationError(
			node,
			`'Group Email Address' should have a maximum length of ${MAX_MAIL_NICKNAME}`,
			{ itemIndex },
		);
	}

	if (mailNickname && !ASCII_MAIL_NICKNAME.test(mailNickname)) {
		throw new NodeOperationError(
			node,
			"'Group Email Address' should only contain characters in the ASCII character set 0 - 127 except the following: @ () \\ [] \" ; : <> , SPACE",
			{ itemIndex },
		);
	}
}
