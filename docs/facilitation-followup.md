# facilitation.followup@1 — follow-ups for native facilitators

Since 0.9.0. Issue
[#18](https://github.com/byte5ai/omadia-m365-connector/issues/18), epic
byte5ai/omadia#2043, kernel spec §5.11.

> **Verification status (2026-11-18).** The three Graph calls were run live in
> the byte5 tenant with an app scoped by RBAC for Applications to one organiser
> mailbox: free/busy for colleagues (F1) and an online-meeting event with join
> link (F2) worked, and a mail was accepted (F3). That is verification **at
> platform level**: the calls were made directly against Graph, not through
> this connector. The connector's own code is verified against a fake Graph
> only (`tests/facilitationFollowup.test.ts`). Still open: how the sender shows
> to a recipient (F3), shared mailboxes (F4), and one end-to-end run through an
> installed connector. See [Manual verification](#manual-verification-on-a-real-instance).

## What it is

A native facilitator holds no business capabilities. After a facilitation
closes it still has to do two things: send the confirmed protocol, and schedule
the follow-up meetings the group agreed on. This service is the only reach it
has into mail and calendars.

| | |
|---|---|
| ServiceRegistry key | `facilitation.followup` |
| Manifest capability | `facilitation.followup@1` |
| Contract | `FacilitationFollowupService`, `FollowupGrant`, `FollowupRefusal` in `@omadia/plugin-api` ≥ 1.55.0; `grantId` and `grant_refused` with the plugin-api of byte5ai/omadia#2119 |
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
- **Free/busy limits.** Graph takes 20 schedules per request and a window
  below 62 days. More than 20 attendees are asked in several requests and the
  answers joined; a window of 62 days or more is refused as `provider_error`
  without a request.
- **Free/busy never touches appointment details.** Graph's answer can
  contain `subject` and `location` of a colleague's appointments, depending on
  how far that person shares their calendar with the organiser (seen live:
  one colleague's items carried both, another's only status and times). The
  service therefore does not read the schedule items at all. It derives the
  busy intervals from `availabilityView` alone, a string with one digit per
  15-minute slot that cannot hold a subject. Nothing but `attendee`, `start`
  and `end` is returned, and nothing of the answer is logged.
- **Granularity.** Busy intervals are exact to 15-minute slots counted from
  `from`, and neighbouring busy slots are one interval, so neither the number
  of appointments nor their exact bounds show. Free and "working elsewhere"
  do not count as busy; tentative, busy, out of office and anything unknown
  do.
- An attendee **missing from `busy`** means Graph could not resolve that
  person's schedule: the schedule came back with an error, without an
  availability view, or not at all
  (typically an external address, or a mailbox the app may not look into).
  That is "availability unknown", not "free", and it does not fail the call;
  an attendee with no appointments is present with `intervals: []`.
- **Invitations** are created with a plain-text body and all attendees as
  required. Creating the event is what sends the invitations. `online: true`
  asks for a Teams meeting (`isOnlineMeeting: true`, `onlineMeetingProvider:
  "teamsForBusiness"`); the link is read from `onlineMeeting.joinUrl`.
- **An online invitation can come back without a link.** With a licensed user
  mailbox as organiser the link came back in the live test. If Graph creates
  the event but returns no `onlineMeeting.joinUrl` (expected, though not
  documented, for an organiser mailbox that cannot organise Teams meetings),
  the result is `{ outcome: 'created', eventId }` **without `joinUrl`**. The
  invitation has been sent and has no Teams link. A caller that asked for
  `online: true` has to check for `joinUrl` and say "invitation created
  without Teams link" instead of reporting plain success. The audit line reads
  `outcome=created_without_join_url`.
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
| `provider_error` | The request is unusable (no addresses, empty subject, bad time window, free/busy window of 62 days or more), or the token request or Graph call failed. The message carries the Graph status and error code. |
| `grant_refused` | Since 0.10.0, checked last, directly before the token and Graph requests. The call states `grant.grantId` and the kernel did not allow it: the grant is unknown, revoked or expired, its execution scope has ended, an operator stopped it, it does not cover the operation, or it does not cover every address (`rejected` then lists them). Also when the kernel has no `grantVerification` service, could not be asked, or gave no usable answer within 10 seconds. The message carries the kernel's code. |

Each call writes one audit line to the plugin log: operation, facilitation id,
number of people, outcome, and what the kernel said about the grant (`grant=`).
Addresses and content are not logged.

### Kernel-issued grants (since 0.10.0)

With kernel grants (byte5ai/omadia#2078) the engine issues a grant under an
execution scope and states its id as `grant.grantId`. The connector then asks
the kernel's `grantVerification` service before every operation, with the
operation (`free_busy`, `create_invitation`, `send_mail`) and every address
the call reaches. The kernel records each question as a use of the grant.

The connector's own checks (allow-list, organiser mailbox, request) run first
and the kernel is asked last, directly before the token and Graph requests. A
call the connector refuses itself never reaches the kernel, so a use the kernel
records as allowed is one that was about to be sent. It can still fail at
Graph; the connector's audit line then shows that outcome next to
`grant=allowed`.

| Situation | What happens | `grant=` in the audit line |
|---|---|---|
| The connector's own checks refuse the call | That refusal; the kernel is not asked | `not_asked:refused_locally` |
| The kernel allows it | The request goes to Graph | `allowed` |
| The kernel refuses | `grant_refused`, no request to Graph | `refused:<kernel code>` |
| The kernel cannot be asked or does not answer | `grant_refused`, no request to Graph | `refused:unavailable` |
| The kernel has no `grantVerification` service | `grant_refused`, no request to Graph: a stated grant has to be checked | `refused:no_kernel_service` |
| The call states no `grantId` | Allow-list only, as before 0.10.0, on any kernel | `not_asked:no_grant_id` |

## What it cannot do

Of a free/busy answer only the availability view is read (see above). The
transport (`src/facilitationFollowup/transport.ts`) takes an operation
name, not a URL, and knows three of them. It has no `GET`, no `PATCH`, no
`DELETE`, and no path into `/messages`, `/mailFolders`, `/calendarView` or a
stored event. The delegated clients (`graphObo.ts`, `graphCalendarClient.ts`,
`/me/...`) are not used: a facilitator acts for no user.

So the service cannot read mail, cannot read calendar contents, and cannot
change or cancel an event, including the ones it created.

## Setup

An administrator has to act before any of this works. A connector app
registration holds no calendar or mail permission unless someone added one; as
of 2026-11-18 neither of byte5's own instances holds `Calendars.*` or
`Mail.Send`.

### 1. Organiser mailbox

Set the setup field **Organisator-Postfach für Facilitation**
(`facilitation_organizer_mailbox`) to the mailbox that sends the protocol and
organises the follow-up meetings. There is no default.

Use a **licensed user mailbox** created for this purpose (for example
`facilitator@…`), not a person's own. Recipients see this mailbox as the
sender, with no "on behalf of" marker, and everything sent is in its Sent
Items and calendar.

A shared mailbox is unproven: the documentation does not cover it for app-only
calls, and a Teams link is expected to need a licence. Use one only after
checklist item F4.

### 2. Which permissions

From the Graph v1.0 reference, "Permissions" table, row "Application":

| Call | Least privileged | Higher privileged |
|---|---|---|
| `calendar: getSchedule` | `Calendars.ReadBasic` (see note) | `Calendars.Read`, `Calendars.ReadWrite` |
| `Create event` | `Calendars.ReadWrite` | — |
| `user: sendMail` | `Mail.Send` | not available |

The working set is **`Calendars.ReadWrite`** and **`Mail.Send`**.
`Calendars.ReadWrite` is needed for the invitation anyway and covers
`getSchedule`.

Note on `Calendars.ReadBasic`: Microsoft's pages disagree (the API page lists
it for applications, the permissions reference only has
`Calendars.ReadBasic.All`, the concept page names `Calendars.Read`), and RBAC
for Applications has no ReadBasic role. `Calendars.ReadWrite` is the only
choice that works in both the Entra and the RBAC model.

### 3. Grant them: RBAC for Applications (current mechanism)

Exchange Online *RBAC for Applications* grants the two roles for the organiser
mailbox only. Exchange Online PowerShell, role Organization Management:

```powershell
Connect-ExchangeOnline

# 1. Register the connector app's service principal in Exchange. Both ids are
#    on the ENTERPRISE APPLICATION, not on the app registration.
New-ServicePrincipal -AppId <app id> -ObjectId <service principal object id> -DisplayName "omadia M365 connector"

# 2. A scope that matches exactly the organiser mailbox (the filter is our example).
New-ManagementScope -Name "omadia facilitation organiser" `
  -RecipientRestrictionFilter "PrimarySmtpAddress -eq 'facilitator@contoso.com'"

# 3. The two roles, limited to that scope.
New-ManagementRoleAssignment -App <app id> -Role "Application Calendars.ReadWrite" `
  -CustomResourceScope "omadia facilitation organiser"
New-ManagementRoleAssignment -App <app id> -Role "Application Mail.Send" `
  -CustomResourceScope "omadia facilitation organiser"

# 4. Check: in scope for the organiser, not for any other mailbox.
Test-ServicePrincipalAuthorization -Identity <app id> -Resource facilitator@contoso.com
Test-ServicePrincipalAuthorization -Identity <app id> -Resource someone.else@contoso.com
```

**Do not grant the same permission in Entra ID as well.** Effective rights are
the union of both, and an Entra grant is tenant-wide ("send mail as any user",
"read and write calendars in all mailboxes"), so it would undo the scope. If
`Calendars.ReadWrite` or `Mail.Send` is already granted to the app in Entra,
remove it there.

Microsoft documents 30 minutes to 2 hours for changes to apply; the test
cmdlet in step 4 bypasses that cache. In the live test the first Graph call
worked 40 seconds after the assignment. No admin consent click is involved.

#### Fallback if scoped free/busy does not work

**Not needed in the tested tenant.** With the app scoped to the organiser
mailbox only, `getSchedule` on that mailbox returned availability for two
colleagues without an error per schedule (F1, byte5 tenant, 2026-11-18); the
same call through a colleague's mailbox was refused with `403
ErrorAccessDenied`, so the scope holds. Microsoft does not document this
behaviour, so another tenant can differ, for example with restricted free/busy
sharing.

Only if F1 shows every colleague as "availability unknown" there: keep the two
assignments above and add a second, read-only one for the participants:

```powershell
New-ManagementScope -Name "omadia facilitation participants" `
  -RecipientRestrictionFilter "<filter matching the participant group>"
New-ManagementRoleAssignment -App <app id> -Role "Application Calendars.Read" `
  -CustomResourceScope "omadia facilitation participants"
```

(An administrative unit can be used as the scope instead.)

### 4. Other ways to grant

**Legacy: application access policy.** Microsoft: application access policies
"are replaced by Role Based Access Control for Applications", and new ones
should not be created. For a tenant that still runs on them: grant
`Calendars.ReadWrite` and `Mail.Send` in Entra ID with admin consent, put the
organiser mailbox into a mail-enabled security group (a shared mailbox can
only be covered through such a group), then

```powershell
New-ApplicationAccessPolicy -AppId <app id> -PolicyScopeGroupId <group> -AccessRight RestrictAccess
```

A policy applies to all Exchange permissions of the app, so check that no
other consumer of this connector app needs another mailbox.

**Unscoped Entra grant (own tenant or evaluation only).** Azure Portal →
**Microsoft Entra ID** → **App registrations** → the connector app → **API
permissions** → **Add a permission** → **Microsoft Graph** → **Application
permissions** → add `Calendars.ReadWrite` and `Mail.Send` → **Grant admin
consent for \<Tenant\>**. Extending an existing registration requires this
renewed admin consent; see *Consent semantics* in
[`INTEGRATION.md`](../INTEGRATION.md) for the `appRoleAssignments` fallback and
the restart. The connector still only calls the organiser mailbox, but the app
could send as anyone and write any calendar.

In every variant the connector app's secret is the credential to protect.

## Manual verification on a real instance

### Owner checklist

Status from the live run in the byte5 tenant on 2026-11-18. "Platform level"
means the Graph call was made directly with an app token scoped by RBAC for
Applications to one licensed user mailbox, not through this connector.

| | Check | Status |
|---|---|---|
| F1 | Scoped free/busy | **Live-verified at platform level.** |
| F2 | Invitation with Teams link | **Live-verified at platform level.** |
| F3 | Mail | **Accepted by Graph (202); sender display not inspected.** |
| F4 | Shared mailbox as organiser | **Open.** |

- **F1 (scoped free/busy).** After the RBAC assignment for the organiser only:
  trigger one follow-up availability check for two colleagues. Expect
  availability for both. Live: `200` with an availability view for each
  person and no error per schedule; the fallback above was not needed.
- **F2 (invitation).** Create one follow-up invitation with Teams link to
  yourself. Expect: the invitation mail from the organiser mailbox, a working
  join link, the event in the organiser's calendar, `joinUrl` in the tool
  result. Live: `201` with `onlineMeeting.joinUrl`, organiser = the mailbox.
  The join link was not opened and the invitation mail was not looked at.
- **F3 (mail).** Send one follow-up mail to yourself. Expect: sender is the
  organiser mailbox without "on behalf of", a copy in the organiser's Sent
  Items. Live: `202 Accepted`. How the sender shows in the inbox and the copy
  in Sent Items were **not inspected**; look at the received mail once.
- **F4 (shared mailbox, optional).** Repeat F2 with a shared mailbox as
  organiser. Expect either a join link or an event without one; that decides
  whether shared mailboxes are allowed. Not attempted.

### Still to do through the connector

None of the above ran through an installed connector 0.9.0. Repeat F1 to F3
once from a real facilitation and confirm alongside:

- With the organiser field empty, a follow-up records `not_configured` and no
  mail or event exists.
- In F1, pick a colleague who shares appointment details with the organiser
  and has an appointment with a recognisable subject: the result has the
  interval and nothing of the subject or location, and neither does the log.
- In F2, check the time zone of the invitation.
- Add an address that is not in the grant to any call: the answer is
  `recipient_not_allowed` naming it, and nobody receives anything from that
  call.
- The plugin log has one `facilitation.followup op=… outcome=…` line per call,
  with no address and no content.

## Sources

Microsoft Learn, read on 2026-11-18; the full source list with page dates is
in the Teams evaluation for byte5ai/omadia#2043.

- Graph v1.0 reference: `calendar: getSchedule`, `Create event`, `user: sendMail`
- Graph: *Get free/busy schedule* (limits: 20 schedules, 62 days), *Online
  meetings in calendar events* and `onlineMeetingInfo` (`joinUrl`), *Send mail
  from another user*
- Graph permissions reference (`Calendars.ReadBasic.All`, `Calendars.ReadWrite`, `Mail.Send`)
- Exchange Online: *Role Based Access Control for Applications* (roles,
  cmdlets, union with Entra grants, cache delay)
- Exchange PowerShell: `New-ApplicationAccessPolicy` (legacy notice)

Not stated on any of these pages: free/busy under a mailbox scope (settled
for the tested tenant by F1), the licence an organiser needs for a Teams link
(F2 worked with a licensed user mailbox; F4 open), and shared mailboxes with
app-only calls (F4 open).
