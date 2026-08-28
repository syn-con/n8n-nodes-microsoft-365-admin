# Upstreaming to n8n

Plan for contributing this package's additions back to n8n's built-in **Microsoft Entra ID**
node (`packages/nodes-base/nodes/Microsoft/Entra`), rather than pursuing community-node
verification.

**Why this route:** verification requires an MIT licence, and this package is a derivative of
n8n's Sustainable Use Licensed Entra node, so it cannot claim MIT. n8n's own verification
guidelines say an iteration on an existing node should be a pull request instead.

## What is already upstream

The service-principal credential we ship locally **already exists in n8n**, including
certificate authentication:

| Upstream | PR | Landed |
|---|---|---|
| `MicrosoftEntraServicePrincipalApi.credentials.ts` | [#32759](https://github.com/n8n-io/n8n/pull/32759) | 2026-06-25 |
| Certificate auth on that credential | [#33420](https://github.com/n8n-io/n8n/pull/33420) | 2026-07-06 |

n8n is also rolling service-principal credentials across the other Microsoft nodes
([#32616](https://github.com/n8n-io/n8n/pull/32616), OneDrive PoC), so this direction has
momentum behind it.

Our local `Microsoft365AdminServicePrincipalApi` is therefore **redundant** — nothing about it
needs upstreaming. What remains is the resources and operations built on top of it.

Note also that upstream `GenericFunctions.ts` already has `handleErrorPostReceive`. Our
error rules are a *refactor* of it, not an addition. n8n does not merge preference
refactors, so each PR below should add its own error-code entries to the existing upstream
function rather than proposing a rewrite.

## Style mismatch to plan around

This package now uses n8n's **programmatic** actions/transport structure: every operation is
a `<operation>.operation.ts` exporting `properties`, `description` and `execute`, dispatched
by `actions/router.ts`.

n8n's built-in Entra ID node is **declarative** — its operations are `routing` blocks on the
parameter definitions, with no `execute` at all. So the operations below cannot be copied
across as they stand; each one has to be translated back into a `routing` block, and the
per-item `execute` bodies re-expressed as `preSend`/`postReceive` hooks.

That mostly affects the mechanical work, not the design: the parameter definitions, the
validation rules and the Graph endpoints all carry over unchanged, and the operation files
are a clearer starting point for a PR than the old single 2,300-line description file was.
The two operations that genuinely cannot be declarative are flagged in the table below.

## The gate (n8n `CONTRIBUTING.md`)

These are hard requirements, not suggestions:

1. **Forum topic first.** "Feature PRs that arrive with no prior discussion will be closed
   with a pointer to the forum." Nothing gets written until n8n agrees the scope.
2. **≤ 1000 added lines per PR**, one logical change each. "Larger or multifaceted PRs will be
   returned for segmentation."
3. **Tests required.** A PR without them is auto-closed after 14 days.
4. **CLA** — one button, [Indie Open Source](https://indieopensource.com/forms/cla) form.
5. **Versioning** — anything that changes existing output must follow the node version
   guidelines. Everything proposed here is purely additive, so this should not apply.
6. **Own words.** "Write the description in your own words. Do not paste raw model output."
   The draft below is a skeleton to rewrite, not to paste.

## Proposed PR sequence

Roughly 3,600 lines of additive surface, which cannot be one PR. Ordered so each lands
independently and the risky parts come last:

| # | Scope | Est. lines | Notes |
|---|---|---|---|
| 0 | Offer the existing `microsoftEntraServicePrincipalApi` credential on the Entra node | ~80 | Node currently declares OAuth2 only. Small, unblocks app-only use of everything below. |
| 1 | **User**: Get Groups, Get Manager, Set Manager, Revoke Sessions | ~450 | Purely declarative. Easiest first real PR. |
| 2 | **Group**: Get Members, Get Owners, Add Owner, Remove Owner | ~550 | Purely declarative. |
| 3 | **Authentication** (read): Get Many Methods, Get Password Method, Delete Method | ~600 | New resource. Needs the `getAuthenticationMethods` list search. |
| 4 | **Authentication** (write): Reset Password, Create Temporary Access Pass | ~700 | Needs `customOperations` — Reset Password must return the generated password, which a 204 PATCH cannot. |
| 5 | **License** (read): Query Tenant Licenses, Query User Licenses, Query License Holders | ~650 | New resource. Needs the `getSubscribedSkus` loader. |
| 6 | **License** (write): Assign, Assign to Group, Unassign | ~800 | Needs `customOperations` to serialise writes — Entra applies one licence change per tenant at a time and rejects concurrent ones. |

`customOperations` is a supported `INodeType` member and is already used in `nodes-base`, so
PRs 4 and 6 do not need new core capability.

**Sequencing risk:** PRs 4 and 6 carry the only non-declarative behaviour. If n8n pushes back
on `customOperations` in this node, PRs 0–3 and 5 still stand on their own.

## Forum topic draft

Post to <https://community.n8n.io/> under Feature Requests. **Rewrite this in your own
voice before posting** — n8n explicitly rejects pasted model output.

---

**Title:** Extending the Microsoft Entra ID node: licence management, authentication methods,
and more user/group operations

I maintain an internal community node that extends n8n's Microsoft Entra ID node for M365
tenant administration. Before opening any PRs I would like to check whether you would accept
this work upstream, and in what shape.

Since [#32759](https://github.com/n8n-io/n8n/pull/32759) and
[#33420](https://github.com/n8n-io/n8n/pull/33420) added the service-principal credential, the
app-only foundation this depends on is already in n8n. The Entra ID node itself still only
declares `microsoftEntraOAuth2Api`, so unattended admin workflows cannot use it yet.

What I would like to contribute, in rough priority order:

1. **Offer the existing service-principal credential on the Entra ID node.** Small change; the
   credential already exists. Everything below is more useful with it.
2. **Licence management** — query tenant SKUs with seat usage, query a user's licences, list
   holders of a SKU, and assign/unassign licences for users and groups. This is the single
   most common M365 admin task I cannot currently automate in n8n. Worth noting the writes
   have to be serialised: Entra applies one licence change per tenant at a time and rejects
   concurrent ones, so this needs `customOperations` rather than plain declarative routing.
3. **Authentication methods** — list a user's registered methods, read the password method,
   delete a method, issue a Temporary Access Pass, and reset a password. Useful for helpdesk
   and offboarding flows. Reset Password also needs `customOperations`, because it has to
   return the generated password and the underlying PATCH answers 204.
4. **More user and group operations** — user: get groups, get/set manager, revoke sessions;
   group: get members, get owners, add/remove owner.

All of it is additive: no existing operation changes shape, so I do not think node versioning
comes into play.

I know the 1000-line limit, so I would split this into roughly seven PRs along the lines
above, each with tests, starting with the credential change and the plain user/group
operations to establish the pattern before the two that need `customOperations`.

Questions before I start:

- Is this direction something you want in the built-in node at all?
- Any preference on ordering, or on splitting differently?
- Is `customOperations` acceptable in this node for the licence writes and Reset Password, or
  would you rather those were handled another way?

Happy to adjust scope. I have the implementation working already, so this is mostly a
question of how you would like it packaged.

---

## Meanwhile

This repo stays on the Sustainable Use License and is not submitted for verification. The
verification-driven fixes already applied (clean scanner run apart from the licence, CI on
`main`, provenance-ready publish workflow) are all worth keeping regardless — they make the
package publishable and the code closer to what upstream review will expect.

Do **not** publish to npm under an MIT claim.
