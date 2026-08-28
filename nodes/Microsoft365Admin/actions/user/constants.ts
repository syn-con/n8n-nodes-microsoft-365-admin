/** The projection behind the "Simplified" output option for a user. */
export const USER_SIMPLE_SELECT =
	'id,createdDateTime,displayName,userPrincipalName,mail,mailNickname,securityIdentifier';

/**
 * The projection behind the "Raw" output option for a user.
 *
 * Graph's own default projection for a user is far narrower than this, so "Raw" names
 * every property the node can read rather than leaving `$select` off.
 */
export const USER_RAW_SELECT =
	'id,accountEnabled,ageGroup,assignedLicenses,assignedPlans,authorizationInfo,businessPhones,city,companyName,consentProvidedForMinor,country,createdDateTime,creationType,customSecurityAttributes,deletedDateTime,department,displayName,employeeHireDate,employeeId,employeeLeaveDateTime,employeeOrgData,employeeType,externalUserState,externalUserStateChangeDateTime,faxNumber,givenName,identities,imAddresses,isManagementRestricted,isResourceAccount,jobTitle,lastPasswordChangeDateTime,legalAgeGroupClassification,licenseAssignmentStates,mail,mailNickname,mobilePhone,officeLocation,onPremisesDistinguishedName,onPremisesDomainName,onPremisesExtensionAttributes,onPremisesImmutableId,onPremisesLastSyncDateTime,onPremisesProvisioningErrors,onPremisesSamAccountName,onPremisesSecurityIdentifier,onPremisesSyncEnabled,onPremisesUserPrincipalName,otherMails,passwordPolicies,passwordProfile,postalCode,preferredDataLocation,preferredLanguage,provisionedPlans,proxyAddresses,securityIdentifier,serviceProvisioningErrors,showInAddressList,signInSessionsValidFromDateTime,state,streetAddress,surname,usageLocation,userPrincipalName,userType';
