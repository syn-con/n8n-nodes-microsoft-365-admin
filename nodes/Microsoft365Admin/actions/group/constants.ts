/** The projection behind the "Simplified" output option for a group. */
export const GROUP_SIMPLE_SELECT =
	'id,createdDateTime,description,displayName,mail,mailEnabled,mailNickname,securityEnabled,securityIdentifier,visibility';

/** The projection used when a group read is asked to expand its members. */
export const GROUP_MEMBER_EXPAND =
	'members($select=id,accountEnabled,createdDateTime,displayName,employeeId,mail,securityIdentifier,userPrincipalName,userType)';
