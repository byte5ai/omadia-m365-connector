# Changelog

Versions before 0.9.1 are described in the commit history and in the
"since x.y.z" notes of `README.md`, `INTEGRATION.md` and `docs/`.

## 0.9.1

Fixes for two defects of the Teams provisioner found at the rollout of native
facilitators (byte5ai/omadia#2043, defects D7 and D8), and a packaging check.
No tenant was available for these changes: they are verified against a fake
Graph, the error body being the one Graph answered at the rollout.

### Fixed

- **A Graph error no longer loses its reason (D7).** A team install refused by
  the team's resource-specific consent check reached the kernel as
  `…backend request [redacted]. Workload Unknown. Request Url:
  https://teamsgraph.teams.microsoft.com/v1.0/groups/c6031502-025…`. Two things
  removed the reason:
  - the secret redaction replaced any run of 60 or more word characters, and
    `IsUserAuthorizedToGrantGroupResourceSpecificPermissionsRequest` is 62
    letters. A run of letters only is now kept; runs with a digit, `-` or `_`
    are redacted as before;
  - error bodies were cut at 200 characters and request URLs at 120. The
    bounds are now 2,000 and 500, after redaction.
- **`ProvisioningRequestError.graphError`** (new, optional): the parsed error
  envelope as fields, `code`, `innerCode`, `message`, `requestId`,
  `clientRequestId` and `body` (at most 4,000 characters). Redacted like the
  message. It holds what the API answered, never a token and never the request
  body. The kernel reads `code`, `innerCode`, `message` and `body`
  (byte5ai/omadia#2095).
- **The chat listing pages to the end (D8).** `listChats` stopped after 20
  pages and returned what it had as if it were everything; on a tenant with
  more than 725 chats the picker's list was short without a sign. It now
  follows `@odata.nextLink` up to 200 pages or 60 seconds.
- **`listChats` and `listTeams` report truncation.** The returned array has a
  non-enumerable `truncated` property: `true` when a page or time bound, or a
  repeating next link, stopped the listing. `listChats` also has
  `withoutMemberNames`, the number of chats that have neither a topic nor
  member names. Iterating, `JSON.stringify` and deep comparison of the array
  are unchanged.

### Changed

- **`listChats` no longer expands members on every page.** Only chats without
  a topic get `memberNames`, through `GET /chats/{id}/members` (delegated
  `Chat.ReadBasic`, the scope the listing already needs): at most 100 per
  listing, group chats first, five at a time. A chat with a topic comes
  without `memberNames`. A consumer that searched chats by member name finds
  only untitled chats that way now.

### Added

- `npm test` runs `npm run package` and checks every entry of the zip against
  a mirror of the kernel's zip ingest rules. `npm test` therefore needs
  `npm run build` first.
