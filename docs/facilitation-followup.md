# facilitation.followup@1 — follow-ups for native facilitators

Since 0.9.0. Issue
[#18](https://github.com/byte5ai/omadia-m365-connector/issues/18), epic
byte5ai/omadia#2043, kernel spec §5.11.

> **Verification status.** There is no test tenant for this feature. The three
> operations are verified against a fake Graph only (`tests/facilitationFollowup.test.ts`),
> shaped after the Graph v1.0 reference. The permissions, the mailbox
> restriction and the consent steps below come from Microsoft's documentation
> and are **not tenant-verified**. Run the [manual checklist](#manual-verification-on-a-real-instance)
> before relying on them.

## What it is

A native facilitator holds no business capabilities. After a facilitation
closes it still has to do two things: send the confirmed protocol, and schedule
the follow-up meetings the group agreed on. This service is the only reach it
has into mail and calendars.

| | |
|---|---|
| ServiceRegistry key | `facilitation.followup` |
| Manifest capability | `facilitation.followup@1` |
| Contract | `FacilitationFollowupService`, `FollowupGrant`, `FollowupRefusal` in `@omadia/plugin-api` ≥ 1.55.0 |
| Consumer | the facilitation engine plugin; it passes the grant with every call |
| Auth | application permissions (client credentials) of the connector app |

It is not listed under `service_types`, so builder-generated agents are not
offered it.

## Operations

Every operation is one `POST` on the configured organiser mailbox.

| Operation | Graph request | Returns |
|---|---|---|
| `freeBusy(grant, { attendees, from, to })` | `POST /users/{organiser}/calendar/getSchedule` | busy intervals per attendee |
| `createInvitation(grant, { attendees, subject, start, end, body, online? })` | `POST /users/{organiser}/events` | `eventId`, and `joinUrl` for an online meeting |
| `sendMail(grant, { recipients, subject, body })` | `POST /users/{organiser}/sendMail` | `sent` |

Details that matter to a consumer:

- **Times** are ISO 8601 instants with an explicit offset (`2026-11-27T10:00:00+01:00`
  or `…Z`). A time without an offset is refused, not guessed. Graph is asked
  and answers in UTC; results are UTC (`…Z`).
- **Free/busy** returns merged intervals only. Of each schedule item just
  `status`, `start` and `end` are read; `subject`, `location` and `isPrivate`
  are never touched, and overlapping appointments are merged so their number
  does not show either. `free` and `workingElsewhere` do not count as busy;
  everything else does.
- An attendee **missing from `busy`** means Graph could not resolve that
  person's schedule (typically an external address). That is "availability
  unknown", not "free"; an attendee with no appointments is present with
  `intervals: []`.
- **Invitations** are created with a plain-text body and all attendees as
  required. Creating the event is what sends the invitations. `online: true`
  asks for a Teams meeting (`isOnlineMeeting`, `onlineMeetingProvider:
  teamsForBusiness`).
- **Mail** is plain text, to `toRecipients`, saved to the organiser's Sent
  Items. Graph answers `202 Accepted`, which means accepted for delivery, not
  delivered.
- There are no retries. A repeated `sendMail` or `createInvitation` would
  reach people twice, so a failure is reported and the decision to repeat stays
  with the caller.

## Refusals

No operation throws; each answers `{ outcome: 'refused', code, message }`.
Checks run in this order, all before any request to Graph or to the token
endpoint:

| Code | When |
|---|---|
| `recipient_not_allowed` | At least one attendee or recipient is not in `grant.allowed`. The whole call is refused; `rejected` lists the addresses. Comparison is case-insensitive on the trimmed address. A string that is not one bare address (`a@x; b@y`, `Name <a@x>`) is rejected even if the grant contains it. |
| `not_configured` | The setup field `facilitation_organizer_mailbox` is empty or not a mailbox address. |
| `provider_error` | The request is unusable (no addresses, empty subject, bad time window), or the token request or Graph call failed. The message carries the Graph status and error code. |

Each call writes one audit line to the plugin log: operation, facilitation id,
number of people, outcome. Addresses and content are not logged.

## What it cannot do

The transport (`src/facilitationFollowup/transport.ts`) takes an operation
name, not a URL, and knows three of them. It has no `GET`, no `PATCH`, no
`DELETE`, and no path into `/messages`, `/mailFolders`, `/calendarView` or a
stored event. The delegated clients (`graphObo.ts`, `graphCalendarClient.ts`,
`/me/...`) are not used: a facilitator acts for no user.

So the service cannot read mail, cannot read calendar contents, and cannot
change or cancel an event, including the ones it created.

## Setup

### 1. Organiser mailbox

Set the setup field **Organisator-Postfach für Facilitation**
(`facilitation_organizer_mailbox`) to the mailbox that sends the protocol and
organises the follow-up meetings. There is no default.

Use a shared or dedicated mailbox, not a person's: everything the facilitator
sends appears in its Sent Items and its calendar. For online meetings the
mailbox needs to be able to organise Teams meetings (licence; from
documentation, not verified).

### 2. Graph application permissions

From the Graph v1.0 reference, "Permissions" table, row "Application":

| Call | Least privileged | Higher privileged |
|---|---|---|
| `calendar: getSchedule` | `Calendars.ReadBasic` (see note) | `Calendars.Read`, `Calendars.ReadWrite` |
| `Create event` | `Calendars.ReadWrite` | — |
| `user: sendMail` | `Mail.Send` | not available |

The set to grant is therefore **`Calendars.ReadWrite`** and **`Mail.Send`**.
`Calendars.ReadWrite` is needed for the invitation anyway and also covers
`getSchedule`, so no third permission is required.

Note on `Calendars.ReadBasic`: the `getSchedule` page lists it for
applications, but the permissions reference has no application permission of
that name, only `Calendars.ReadBasic.All`. Since `Calendars.ReadWrite` is
required regardless, this discrepancy does not affect the set above.

Both permissions are tenant-wide when granted in Entra ID: "send mail as any
user" and "read and write calendars in all mailboxes". The connector only ever
addresses the organiser mailbox, but the token could do more. Restrict it:

### 3. Least privilege, by tenant situation

**A. Exchange Online with an Exchange administrator available (recommended).**
Use *RBAC for Applications in Exchange Online* and scope both roles to the
organiser mailbox. In this option the two permissions are **not** granted in
Entra ID at all: Microsoft documents the Exchange grants as independent of the
unscoped Entra grants, and states that an existing organisation-wide grant of
the same permission has to be removed for the scope to take effect.

```powershell
Connect-ExchangeOnline

# The connector app's service principal: AppId = App (Client) ID,
# ObjectId = object id of the ENTERPRISE APPLICATION (not the app registration).
New-ServicePrincipal -AppId <app-id> -ObjectId <enterprise-app-object-id> -DisplayName "omadia M365 connector"

# Scope = exactly the organiser mailbox. The filter is our example.
New-ManagementScope -Name "omadia facilitation organiser" `
  -RecipientRestrictionFilter "PrimarySmtpAddress -eq 'facilitator@contoso.com'"

New-ManagementRoleAssignment -App <app-id> -Role "Application Calendars.ReadWrite" `
  -CustomResourceScope "omadia facilitation organiser"
New-ManagementRoleAssignment -App <app-id> -Role "Application Mail.Send" `
  -CustomResourceScope "omadia facilitation organiser"

# Expect InScope = True for the organiser, False for any other mailbox.
Test-ServicePrincipalAuthorization -Identity <app-id> -Resource facilitator@contoso.com
Test-ServicePrincipalAuthorization -Identity <app-id> -Resource someone.else@contoso.com
```

Changes take effect after 30 minutes to 2 hours (permission cache); the test
cmdlet bypasses the cache. No admin consent click is involved in this option.

**B. Tenants that already use application access policies (legacy).**
Grant `Calendars.ReadWrite` and `Mail.Send` in Entra ID with admin consent (as
in C), then restrict the app to a mail-enabled security group that contains
only the organiser mailbox:

```powershell
New-ApplicationAccessPolicy -AppId <app-id> -PolicyScopeGroupId <group> `
  -AccessRight RestrictAccess -Description "omadia facilitation follow-ups"
Test-ApplicationAccessPolicy -Identity facilitator@contoso.com -AppId <app-id>
```

Microsoft states that RBAC for Applications replaces application access
policies, and its former policy page now leads to the RBAC article. The two
cmdlets above are written from the earlier documentation and were not
re-checked against a current page. Prefer A for a new setup. A policy applies
to all Exchange permissions of the app, so check that no other consumer of
this connector app needs another mailbox.

**C. Unscoped (own tenant, evaluation, or no Exchange administrator).**

1. Azure Portal → **Microsoft Entra ID** → **App registrations** → the
   connector app → **API permissions** → **Add a permission** → **Microsoft
   Graph** → **Application permissions**.
2. Add `Calendars.ReadWrite` and `Mail.Send`.
3. **Grant admin consent for \<Tenant\>** (needs an administrator). Extending
   an existing registration requires this renewed admin consent; see
   *Consent semantics* in [`INTEGRATION.md`](../INTEGRATION.md) for the
   `appRoleAssignments` fallback and the restart.

This works, and the connector still only calls the organiser mailbox, but the
app could send as anyone and write any calendar. Treat it as a stopgap in a
customer tenant.

In all three options the connector app's secret is the credential to protect.

## Manual verification on a real instance

Nothing below has been run against a tenant yet. Use a facilitation whose
grant lists two internal test users, and one address that is not in it.

1. **Not configured.** Leave the organiser field empty, trigger a follow-up:
   the protocol records `not_configured`, and no mail or event exists.
2. **Free/busy.** Give one test user an appointment with a recognisable
   subject in the window. `freeBusy` returns the interval for that user and
   nothing of the subject or location; the other user appears with
   `intervals: []`.
3. **Free/busy under option A or B.** Repeat step 2 with the mailbox
   restriction active. This is the open question: `getSchedule` runs on the
   organiser's calendar but returns other people's availability, and whether a
   scope limited to the organiser mailbox still allows that is not documented.
   If it answers 403 or an `error` per schedule, the scope has to include the
   participants for `Calendars.ReadWrite`, or free/busy stays unavailable.
4. **Invitation.** `createInvitation` for the two users: both receive the
   invitation from the organiser mailbox with subject, time (check the time
   zone) and agenda; the event is in the organiser's calendar.
5. **Online meeting.** With `online: true` the result carries a `joinUrl` and
   the invitation contains the Teams link. Without a `joinUrl`, check the
   organiser mailbox's Teams licence.
6. **Mail.** `sendMail` to the two users: both receive the plain-text protocol
   from the organiser mailbox; it is in Sent Items.
7. **Outside the grant.** Add the third address to any of the three calls:
   the answer is `recipient_not_allowed` naming it, and neither of the granted
   users receives anything from that call.
8. **Restriction holds.** Under A or B, `Test-ServicePrincipalAuthorization`
   (or `Test-ApplicationAccessPolicy`) denies a mailbox other than the
   organiser.
9. **Log.** The plugin log has one `facilitation.followup op=… outcome=…` line
   per call, with no address and no content.

## Sources

Read on 2026-11-18:

- Graph v1.0 reference: `calendar: getSchedule`, `Create event`, `user: sendMail`
- Graph permissions reference (`Calendars.ReadBasic.All`, `Calendars.ReadWrite`, `Mail.Send`)
- Exchange Online: *Role Based Access Control for Applications*
  (roles, cmdlets, relation to Entra grants, cache delay)

The `isOnlineMeeting` / `onlineMeeting.joinUrl` shape and the application
access policy cmdlets are from earlier knowledge of those pages, not from a
verbatim read on that date.
