# Microsoft 365 Admin

An n8n node for managing users, groups, authentication methods and licenses in Microsoft
Entra ID through the Microsoft Graph API.

It authenticates as an application rather than a user, so it runs unattended: no browser
sign-in, no redirect URI, and no refresh token to expire.

## Operations

**User** — Create · Get · Get Many · Update · Delete · Add to Group · Remove from Group ·
Get Groups · Get Manager · Set Manager · Revoke Sessions

**Group** — Create · Get · Get Many · Update · Delete · Get Members · Get Owners ·
Add Owner · Remove Owner

**Authentication** — Get Many Methods · Get Password Method · Delete Method ·
Create Temporary Access Pass · Reset Password

**License** — Query Tenant Licenses · Query User Licenses · Query License Holders ·
Assign · Assign to Group · Unassign

## Install

The package is published to GitHub Packages, so npm needs to be pointed at that registry
for the `@syn-con` scope. In `~/.npmrc`:

```
@syn-con:registry=https://npm.pkg.github.com
//npm.pkg.github.com/:_authToken=YOUR_GITHUB_TOKEN
```

n8n 1.81 or newer is required: the license write operations run as custom operations, which
older versions do not execute.

Then install it into your n8n instance and restart:

```bash
cd ~/.n8n/nodes
npm install @syn-con/n8n-nodes-microsoft-365-admin
```

## Setup

### 1. Register an application

In the [Entra admin center](https://entra.microsoft.com) go to **Applications** →
**App registrations** → **New registration**. Choose **Single tenant** and leave the
redirect URI blank — this node never uses one.

From the **Overview** page copy the **Application (client) ID** and the
**Directory (tenant) ID**. Under **Certificates & secrets** create a client secret and
copy its **Value** straight away; it is only shown once.

### 2. Grant permissions

Under **API permissions** → **Add a permission** → **Microsoft Graph**, choose
**Application permissions**. Not Delegated — delegated permissions do nothing for an
app-only connection and every call comes back 403.

| Permission | Covers |
|---|---|
| `Organization.Read.All` | the connection test and tenant license queries |
| `User.ReadWrite.All` | all user operations and user licensing |
| `Group.ReadWrite.All` | all group operations, including members and owners |
| `UserAuthenticationMethod.ReadWrite.All` | Authentication → Get Many Methods, Get Password Method, Delete Method and Create Temporary Access Pass |
| `User-PasswordProfile.ReadWrite.All` | Authentication → Reset Password through `passwordProfile` |

Then click **Grant admin consent**. This needs Global Administrator or Privileged Role
Administrator — Application Administrator is not enough.

Reset Password additionally requires the application's service principal to hold at least
the **User Administrator** Microsoft Entra role. Resetting privileged administrators is
subject to the Entra role hierarchy and can require **Privileged Authentication
Administrator**. These directory roles are separate from Microsoft Graph application
permissions.

### 3. Add the credential

In n8n create a **Microsoft 365 Admin (Service Principal) API** credential and fill in the
tenant ID, client ID and secret. Leave the Graph base URL on **Global** unless your tenant
is in a sovereign cloud.

Save to run the connection test.

## Things worth knowing

**Recreate the credential after changing permissions.** The access token is cached on the
credential. Graph replies 403 to a valid token that lacks a permission, and n8n only
fetches a new token on a 401 — so a token minted before you granted consent keeps failing
forever. Deleting and re-adding the credential is the fix; editing and saving is not.

**Entra ID applies one license change per tenant at a time.** A second `assignLicense`
write arriving while the first is still being processed is rejected outright — "Error due
to concurrent requests being made to the tenant" — and the tenant can stay busy for the
better part of a minute. Assign, Assign to Group and Unassign therefore behave differently
from every other operation in this node: they send one request at a time rather than one
per input item in parallel, and a rejected write is retried with a growing delay instead of
failing the run. Three things make a bulk change finish quickly:

- **Pick every SKU in one item.** The License SKU field is a multi-select, and one request
  carries any number of licenses for the same tenant-processing time. Ten licenses on one
  user is one request, not ten.
- **Swap in place.** Options → *License SKU Names or IDs to Remove* on Assign removes
  licenses in the same request that adds the new ones, so an E3 → E5 move costs one round
  of processing instead of two.
- **Let items collapse.** Options → *Combine Items for the Same Target* (on by default)
  merges every input item aimed at the same user into a single request, so a workflow that
  emits one item per license change still sends one request per user.

For hundreds of users, license through a group instead: **Assign to Group** is a single
write that Entra fans out to the members itself, in the background.

Options also carries *Max Retries* (default 5) and *Wait Between Requests*, for tenants
where something outside the workflow — an admin in the portal, another automation — is
competing for the same lock.

**Password reset goes through `passwordProfile`, not the authentication methods API.**
Graph's `POST /users/{id}/authentication/methods/{id}/resetPassword` supports delegated
access only — there is no application permission for it, so an app-only credential like
this one can never call it, whatever roles the app holds. Authentication → **Reset Password**
therefore PATCHes **passwordProfile** instead, which is the app-only equivalent: it writes
to Entra ID and, where password writeback is configured, on to on-premises AD. Leave the
Password option empty and the node generates one and returns it as `password` on the output
item — the only place a generated password exists, so capture it in the same run.

**Deleting a method needs its type as well as its ID.** Each method type lives in its own
Graph collection, and `Get Many Methods` reports an `@odata.type` rather than the URL
segment a delete needs. The node fills that gap: every method it returns carries
`methodType` (`phoneMethods`, `fido2Methods`, …), `methodName` and `deletable`, so an MFA
reset is Get Many Methods → Filter on `deletable` → Delete Method with
`{{ $json.methodType }}` and `{{ $json.id }}`. Picking by hand works too: choose the type,
and the Method list shows what that user has registered.

**Some methods refuse to be deleted.** Graph rejects deleting a phone number that is the
user's default MFA method — the user has to change their default first — and there is no
delete for a password at all. A `mobile` number cannot be removed while an
`alternateMobile` one exists.

**A Temporary Access Pass is returned once.** The `temporaryAccessPass` field of the
response is the passcode; it is never readable again. A user can hold only one usable pass,
so issuing a second replaces the first. A multi-use pass is only accepted if the Temporary
Access Pass authentication method policy allows it, and the method has to be enabled in
that policy for the target user at all — otherwise Graph refuses to issue one.

**Revoking sessions is under User, not Authentication** — User → Revoke Sessions
invalidates refresh tokens and browser sessions, which pairs well with a password reset.

**Set `usageLocation` before licensing a user.** Assigning a license to a user without one
fails with an error that looks like a permissions problem. Set it first with User → Update
using a two-letter country code, e.g. `LT`.

**Licenses inherited from a group cannot be unassigned directly.** Query License Holders
returns `licenseAssignmentStates`; if `assignedByGroup` is set, the license comes from
group-based licensing and Unassign silently does nothing. Remove the user from that group
instead.

**Assign to Group needs Entra ID P1** and licenses members in the background, so a
successful response does not mean they are licensed yet.

**Users holding admin roles need more than Graph permissions.** To modify them the app
also has to be assigned a directory role such as User Administrator — and Authentication
Administrator for the Authentication resource, or Privileged Authentication Administrator
to act on privileged accounts.

## Using expressions

Any field can be driven from workflow data. Resource pickers (User, Group, Manager, Owner)
take an expression once you switch them to **By ID**; dropdowns such as License SKU take
one via the Fixed/Expression toggle.

**Resource** and **Operation** are the exceptions — n8n needs to know the operation to
decide which fields to show. Use a Switch node if you need to vary it at runtime.

## Development

```bash
npm install
npm run dev     # n8n with this node linked, on http://localhost:5678
npm run build
npm run lint
npm test
```

`npm run test:coverage` enforces an 80% floor on lines, statements, branches and functions.

## License

Derived from the Microsoft Entra ID node in [n8n](https://github.com/n8n-io/n8n) and
distributed under n8n's Sustainable Use License — **not** MIT. Internal business use is
permitted; redistribution is not, other than free of charge for non-commercial purposes.
See [LICENSE.md](LICENSE.md).
