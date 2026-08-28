import type { AllEntities } from 'n8n-workflow';

type NodeMap = {
	authentication:
		| 'createTemporaryAccessPass'
		| 'deleteMethod'
		| 'getAllMethods'
		| 'getPasswordMethod'
		| 'resetPassword';
	group:
		| 'addOwner'
		| 'create'
		| 'delete'
		| 'get'
		| 'getAll'
		| 'getMembers'
		| 'getOwners'
		| 'removeOwner'
		| 'update';
	license: 'assign' | 'assignGroup' | 'queryHolders' | 'queryTenant' | 'queryUser' | 'unassign';
	user:
		| 'addGroup'
		| 'create'
		| 'delete'
		| 'get'
		| 'getAll'
		| 'getGroups'
		| 'getManager'
		| 'removeGroup'
		| 'revokeSessions'
		| 'setManager'
		| 'update';
};

export type Microsoft365Admin = AllEntities<NodeMap>;
