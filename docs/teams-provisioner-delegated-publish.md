# Delegated catalog publishing

Since **0.6.0** — [byte5ai/omadia#924](https://github.com/byte5ai/omadia/issues/924).

## The problem

Every step of the Teams provisioning chain runs app-only with client
credentials, except one. `POST /appCatalogs/teamsApps` — the upload of the
generated app package into the tenant app catalog — is **delegated-only**.

Microsoft Graph's own reference for
[Publish teamsApp](https://learn.microsoft.com/en-us/graph/api/teamsapp-publish)
states it plainly:

| Permission type | Least privileged | Higher privileged |
| --- | --- | --- |
| Delegated (work or school account) | `AppCatalog.Submit` | `AppCatalog.ReadWrite.All`, `Directory.ReadWrite.All` |
| Delegated (personal Microsoft account) | Not supported. | Not supported. |
| **Application** | **Not supported.** | **Not supported.** |

Confirmed in the field: with one app-only token, the catalog **lookup**
(`GET /appCatalogs/teamsApps?$filter=externalId eq '…'`) succeeds and the
**upload** is refused — while `AppCatalog.ReadWrite.All` is assigned as an app
role and admin-consented. It is not a consent problem, so no amount of consent
fixes it.

Note also that `AppCatalog.Submit`, listed as least-privileged, is **not
sufficient**: Graph's own note says it "allows you to submit apps for review
only, not to publish them to the catalog". Publishing needs delegated
`AppCatalog.ReadWrite.All`, which is admin-consent-required.

## The shape of the solution

One tenant admin signs in **once**. Every agent provisioned afterwards runs
fully automatically. No manual package upload, ever.

```
startDelegatedSignIn ─┐
                      ├─▶ (admin opens the URL, types the code, consents)
pollDelegatedSignIn ──┘        │
                               ▼
                     DelegatedTokenSet  ──── caller persists it ────┐
                               │                                    │
                               ▼                                    │
                     uploadToCatalogDelegated ◀────────────────────-┘
                               │
                               ├─▶ refreshes the access token when stale
                               └─▶ returns a possibly-ROTATED token set
```

### Why the device authorization grant, not authorization code

An authorization-code (redirect) flow needs a redirect URI registered in Entra
**per deployment URL**. omadia is self-hosted: every instance answers on a
different host, and moving one would silently break the flow until someone
edited the app registration by hand. The device authorization grant
([RFC 8628](https://tools.ietf.org/html/rfc8628)) has no redirect URI at all,
so the same publisher app works for every install forever.

### Why raw `fetch`, not MSAL

`@azure/msal-node` is a dependency (see `src/graphObo.ts`), but
`acquireTokenByDeviceCode` **blocks** until the user finishes and reports the
user code through a callback. A device-code flow inherently spans two HTTP
requests of the operator — "show me the code" and, minutes later, "am I done
yet?" — so a blocking call would force the connector to hold flow state in a
process-local map. That map dies on every deploy and does not exist on a second
instance, which turns a restart during sign-in into a silent dead end.

Speaking to `/devicecode` and `/token` directly makes the flow **pollable from
any process**: all state lives in the handle the caller holds. It is also the
house style — `AadTokenCache` in `graphClient.ts` already speaks the same
endpoint with hand-rolled form posts — and it keeps msal confined to the OBO
path.

## The publisher app

The device code grant only works for a **public client**, i.e. an app Entra
issues tokens to without a client secret. The connector's own app registration
must never become one: it holds `Application.ReadWrite.OwnedBy`, so it can mint
app registrations and secrets across the tenant. An identity with that reach and
no secret required to speak as it is not a trade worth making for one upload.

So the connector registers a **separate, minimal app** through Graph, using the
permission it already has:

| Property | Value | Why |
| --- | --- | --- |
| `signInAudience` | `AzureADMyOrg` | The provisioner's SingleTenant invariant. |
| `isFallbackPublicClient` | `true` | "Allow public client flows" — required by the device code grant. |
| `publicClient.redirectUris` | `[]` | Deliberately empty; nothing to register per deployment. |
| `requiredResourceAccess` | `AppCatalog.ReadWrite.All`, `type: "Scope"` | Exactly one **delegated** permission. `Role` would be an app permission, which Graph refuses for this verb. |
| password credentials | *none, ever* | It borrows the signed-in admin's rights; it can do nothing on its own. |
| `uniqueName` | `omadia-teams-publisher-<tenantId>` | Idempotency key; stable across reinstalls. |

It is created **lazily**, on the first `startDelegatedSignIn` — an install that
never publishes a Teams app never registers one, and activation stays
side-effect free.

Idempotency reuses the app-registration step's machinery from
[byte5ai/omadia#916](https://github.com/byte5ai/omadia/issues/916) rather than
re-deriving it: a taken `uniqueName` is adopted, Entra's replication windows are
polled through, and a name held by a soft-deleted app is recovered by restoring
that app. An adopted app is also **reconciled** — one registered by an older
version, or edited in the portal, may be missing the public-client flag or the
scope, and both are silent failures at sign-in time.

**There is no rollback, deliberately.** A delete soft-deletes the object and
reserves its `uniqueName` for 30 days, so a rollback on a half-created publisher
app would lock the tenant out of signing in for a month. There is also nothing
worth undoing: the app holds no secret and grants nobody anything until an admin
consents.

## Consent

`AppCatalog.ReadWrite.All` (delegated) has `AdminConsentRequired: Yes`
(permission id `1ca167d5-1655-44a1-8adf-1414072e1ef9`).

**`prompt=consent` is not available here.** The `/devicecode` endpoint documents
exactly two parameters, `client_id` and `scope` — there is no `prompt`. Consent
therefore has to happen in one of two places:

1. **At the sign-in itself.** An administrator opening the verification URL for
   an app the tenant has not consented to is shown the consent screen and can
   grant on behalf of the organization. This is the normal path and needs no
   extra step.
2. **Ahead of time, through the admin-consent URL** — needed when the tenant's
   user-consent policy is "do not allow user consent", or when the person
   signing in is not an administrator. Every relevant surface carries it:
   `DeviceCodeStart.adminConsentUrl`, `DelegatedConsentRequiredError.adminConsentUrl`,
   and `DelegatedSignedInStatus.adminConsentUrl`.

   ```
   https://login.microsoftonline.com/<tenant-id>/adminconsent?client_id=<publisher-app-id>
   ```

   It contains only the public client id — never a secret.

## Known risk: Conditional Access can block this flow

Microsoft's guidance is explicit: *"Device code flow is a high-risk
authentication method… Allow device code flow only where necessary. Microsoft
recommends blocking device code flow wherever possible."* Entra has a dedicated
Conditional Access **authentication flows** condition for exactly this, and it
is a common hardening baseline.

In such a tenant the browser leg is refused and the poll returns
`{ status: 'declined' }` — the same discriminant as an admin clicking cancel.
The `reason` field carries Entra's redacted `error_description` including its
AADSTS code, which is the only thing that tells the two apart. **Read it before
blaming the operator.**

Note also Entra's *protocol tracking*: a session established through device code
flow stays marked, and later refreshes on that session remain subject to
authentication-flow policies.

**If a customer tenant blocks device code flow, this design cannot work there.**
The fallbacks, in order of preference:

1. Ask the tenant admin to scope the Conditional Access policy so it excludes
   the publisher app (the recommended narrow exception — the app holds one
   delegated scope and no secret).
2. Fall back to an authorization-code flow with a redirect URI registered per
   deployment. This is what the device code grant was chosen to avoid, but it is
   the only other flow that yields a delegated token without a human at a
   terminal.
3. Manual package upload per agent — the outcome this work exists to eliminate.

## Token lifecycle

- `offline_access` is requested, so the sign-in yields a **refresh token**.
  Without one the admin would have to sign in every hour, defeating the design;
  a token response without one is therefore a hard failure at sign-in time
  rather than a problem discovered an hour later.
- Entra **rotates** the refresh token on every exchange. `uploadToCatalogDelegated`
  and `refreshDelegatedToken` return the rotated set with a `refreshed` flag —
  a caller that ignores it and keeps writing back the original eventually forces
  a pointless second admin sign-in.
- The **caller owns persistence.** A plugin has no database, so the connector
  takes credentials as parameters and hands renewed ones back. It caches
  nothing and writes nothing.

## Error taxonomy

Three failure modes, three classes, because each has a different remedy:

| Error | Meaning | What the operator must do |
| --- | --- | --- |
| `DelegatedSignInRequiredError` | No delegated token was supplied. | Start the device-code sign-in. |
| `DelegatedConsentRequiredError` | Signed in, but the tenant has not consented. | Send an admin to `adminConsentUrl`. |
| `DelegatedTokenExpiredError` | The stored credential stopped working. | `reason: 'access-token-expired'` → refresh (no human). `'refresh-token-invalid'` → sign in again. |
| `DeviceCodeFlowError` | The protocol itself failed. | Read `oauthError`; check the publisher app and any Conditional Access policy. |

`authorization_pending`, `expired_token` and `authorization_declined` are **not**
errors — they are the poll result's discriminant, so a caller renders them
instead of catching them.

## Secrets

Four values must never reach a log line, an error message, an error `cause`
chain, or a returned field: the delegated **access token**, the **refresh
token**, the **device code**, and the app **client secret**. The `user_code` is
the exception — it exists to be displayed.

`src/teamsProvisioner/redact.ts` is the last line of defence, scrubbing JWTs,
bearer headers, named credential fields in both JSON and form encodings, and
bare long opaque tokens from anything on the error path — including bodies the
connector did not write. `tests/provisionerDelegatedRedaction.test.ts` drives
the real code paths with sentinel credentials and asserts on every log line,
every message in the cause chain, and every field of every result.

The **flow handle** returned by `startDelegatedSignIn` is documented as
**secret-grade**: it carries the RFC 8628 `device_code`, and whoever holds that
can redeem the token the admin is about to sign for. It is base64url-encoded
JSON — framing, not encryption. Store it like a password, never render it, and
drop it once the flow reaches a terminal state. Its ~15-minute protocol lifetime
bounds the damage; the encoding does not protect it.
