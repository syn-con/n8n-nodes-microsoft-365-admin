# Changelog

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
- Requires n8n 1.80 or newer
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

