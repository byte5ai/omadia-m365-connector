# Changelog

Versions before 0.9.1 are described in the commit history and in the
"since x.y.z" notes of `README.md`, `INTEGRATION.md` and `docs/`.

## 0.10.0

The connector verifies the kernel-issued grant (byte5ai/omadia#2078, spec 2043
§5.11, M1 of the spec alignment). Verified against a fake kernel service and a
fake Graph; nothing ran against a kernel or a tenant.

### Added

- **`facilitation.followup` asks the kernel before every operation.** When the
  caller's `FollowupGrant` states `grantId`, `freeBusy`, `createInvitation` and
  `sendMail` first call `grantVerification.verify({ grantId, operation,
  recipients })` with `free_busy`, `create_invitation` or `send_mail` and every
  address the call names (trimmed, lower-cased, de-duplicated). The kernel is
  asked at every call, never cached. `freeBusy` asks once for all attendees,
  not once per getSchedule request.
- **New refusal code `grant_refused`.** Any answer other than `allowed: true`
  ends the call before the token endpoint and Graph are reached. The message
  carries the kernel's code and wording, for example
  `the kernel refused the grant (revoked): this grant was revoked`; for
  `recipient_not_covered`, `rejected` lists the addresses the kernel named.
- **No answer refuses.** The kernel answering `unavailable`, `verify()`
  throwing, no answer within 10 seconds, an answer of another shape, a failing
  service lookup, or a service that was there and is gone: all `grant_refused`.
- **The audit line states the verification.** A new last field `grant=` with
  `allowed`, `refused:<kernel code>`, `refused:unavailable`,
  `not_asked:no_grant_id` or `not_asked:no_kernel_service`.
- Manifest: `optional_requires: ["grantVerification@^1"]`. The service is
  looked up at every call, because the kernel publishes it after this plugin
  activates.

### Unchanged

- **The allow-list of the call stays the second check.** After the kernel
  allowed the operation, every address still has to be in `grant.allowed`,
  else `recipient_not_allowed` as before.
- **A caller without `grantId`** (an engine before kernel grants) is not
  verified with the kernel: allow-list only, as in 0.9.x.
- **A kernel without `grantVerification`** (before #2078): allow-list only,
  also when the caller states a grant id. This is logged once per activation.

### Changed

- Every audit line of `facilitation.followup` ends with the new `grant=` field.
  Anything that parses these lines to the end of the line has to allow for it.

## 0.9.3

Fix for rollout defect D22 of native facilitators (byte5ai/omadia#2043).
Verified against a fake Graph; nothing ran against a tenant.

### Fixed

- **A 403 on a follow-up write to a fresh app registration is no longer
  reported as missing admin consent.** `applications.addPassword` and
  `servicePrincipals.create` run seconds after the registration was created or
  found with `Application.ReadWrite.OwnedBy`, so the permission is granted; a
  403 there is Entra not having replicated the new object's owner yet. It is
  now retried on the same budget as the 404 (8 attempts, about 40 s). Live, the
  403 came 8 seconds after the create and the identical re-run succeeded.
- When that budget runs out the error is the transient
  `DirectoryReplicationError`, not `ConsentMissingError`. It has two new
  optional fields, `lastStatus` (403 or 404) and `lastDetail` (Graph's
  wording), and for a 403 its message says so:
  `directory_replication_pending: graph <step> still answered 403 after …`.
  The message of the 404 case is unchanged.
- The refused call and Graph's answer are logged at each such retry.

### Unchanged

- A 403 on `applications.create` or `applications.findByUniqueName` is still
  `ConsentMissingError`, on the first answer.

## 0.9.2

Fixes for rollout defect D10 of native facilitators (byte5ai/omadia#2043): after
a reset, the next provisioning produced a bot the Bot Connector rejects with 401
for good, while Azure reported it healthy. The cause was a chain: a purge asked
milliseconds after the delete answered `already-absent`, the next run restored
the soft-deleted application with its old client id, and the bot resource was
re-created under that app id. Verified against a stateful fake of Entra, Azure
Resource Manager and the Bot Connector; nothing ran against a tenant.

### Added

- **`probeBot({ appId })`**: asks the Bot Connector, as the bot, whether it knows
  the bot. `rejected` for a Connector 401 or an `invalid_client` at the token
  endpoint, `known` for any 2xx or a 400/403/404/405, `unverified` otherwise.
  The kernel calls it when present (byte5ai/omadia#2097). Not verified against
  a tenant: what a healthy bot answers for the fabricated member id the probe
  sends; details in `docs/teams-provisioner.md`.
- **`findDeletedAppRegistration({ appId })`**: looks for an application in the
  recycle bin. Throws `RecycleBinUnreadableError` when the bin cannot be read.
  The kernel's reset confirms a purge with it; without it that confirmation
  was skipped.

### Fixed

- **A purge no longer reports `already-absent` for an application deleted
  moments ago.** After a delete made through the connector it looks into the
  recycle bin up to 6 times (about 30 s), purges the tombstone when it appears,
  and answers the new outcome `unconfirmed` (with `detail`) when it does not.
- **`createAppRegistration` never restores a soft-deleted application.** A
  tombstone holding the `uniqueName` is permanently deleted and a new
  application is created, with a new client id.
- **`createBot` refuses an app id whose bot resource was deleted**
  (`BotAppIdRetiredError`, `bot_app_id_retired: …`). `deleteBot` records the
  app id in the plugin vault under `teams_retired_bot_app:<appId>`.
- **A failed Teams-channel enablement keeps the bot resource** instead of
  deleting it, so the re-run does not create a second one under the same app id.

### Changed

- `purgeDeletedAppRegistration` answers `unconfirmed` instead of
  `already-absent` when the recycle bin cannot be read.
- `deleteAppRegistration` and `deleteBot` each make one read before the delete
  (object id, app id).
- A bot deleted before 0.9.2 is not recorded as retired; the guard covers
  deletes made from this version on.

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
