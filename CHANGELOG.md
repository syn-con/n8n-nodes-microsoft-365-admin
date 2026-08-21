# Changelog

## 0.3.1

- Empty method pickers now consistently show No results. They filter the aggregate methods
  endpoint instead of surfacing the 404 that Graph returns for an empty Platform Credential
  collection.

## 0.3.0

- New **Authentication** resource for password and MFA administration: Get Many Methods,
  Get Password Method, Delete Method, Create Temporary Access Pass and Reset Password
- Get Many Methods adds `methodType`, `methodName` and `deletable` to each method, so its
  output can be piped straight into Delete Method — Graph reports an `@odata.type`, not the
  collection segment the delete URL needs
- Delete Method covers every type Graph can delete in v1.0 (Microsoft Authenticator, phone,
  FIDO2, email, software OATH, Temporary Access Pass, Windows Hello, platform credential)
  and its Method picker lists what the chosen user actually has registered
- Reset Password writes `passwordProfile`, generates a password when none is given and
  returns it on the output item. Graph's `authentication/methods/{id}/resetPassword` is
  delegated-only, so an app-only credential cannot call it.
- Create Temporary Access Pass now sends an empty JSON object when no options are selected,
  allowing Graph to apply the tenant policy defaults instead of rejecting an empty payload
- Setup guidance now includes the new `UserAuthenticationMethod.ReadWrite.All` and
  `User-PasswordProfile.ReadWrite.All` application permissions. Reset Password also
  documents its required User Administrator service-principal role.
- Revoking sessions was already available as User → Revoke Sessions
- Internal: Graph error handling moved out of GenericFunctions into GraphErrors, the User
  picker and the custom-operation error helpers are now shared

## 0.2.0

- License Assign, Assign to Group and Unassign now send their requests one at a time and
  retry the "concurrent requests being made to the tenant" rejection with backoff, instead
  of firing one request per input item in parallel and failing most of them
- The License SKU field on those operations takes several SKUs at once, so any number of
  license changes for one user costs a single Graph request
- Assign can remove licenses in the same request that adds them (Options → License SKU
  Names or IDs to Remove), turning a swap into one round of tenant processing
- Items aimed at the same user or group are merged into one request by default
  (Options → Combine Items for the Same Target)
- Disabled Plans is now sent only with the SKU that contains each plan, and a plan that
  belongs to none of the selected SKUs is reported instead of being sent to Graph
- A user or group ID carrying URL characters is now rejected instead of being pasted into
  the request path, and each item of a merged request gets its own copy of the response
- Searching the User or Group picker for a name containing an apostrophe or a quote no
  longer fails — the term is escaped before it goes into the Graph query
- The node and credential ship a dark-theme icon variant
- Internal: one Graph request builder shared by every operation, the Graph error messages
  moved to a rule table, and the user field conversions shared by create and update — the
  package now lints clean with no warnings
- Requires n8n 1.81 or newer
- Existing workflows keep working — a single SKU ID, and a comma-separated list from an
  expression, are both still accepted — but the field is now a multi-select, so re-open the
  node if the picker looks empty

## 0.1.1 – 0.1.3

- Release plumbing only: publish workflow and package metadata. The node itself is
  unchanged from 0.1.0.

## 0.1.0

- Initial release: Microsoft 365 Admin node with User, Group and License resources
- App-only (service principal) authentication via the client credentials grant
- Derived from the built-in n8n Microsoft Entra ID node; see LICENSE.md
