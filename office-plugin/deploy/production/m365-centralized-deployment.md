# Microsoft 365 Centralized Deployment

Use this once the hosted add-in service is reachable over HTTPS.

## Inputs you need

- Hosted add-in origin, for example `https://ctrl-ai.expedient.cloud`.
- Production manifest generated with that exact origin:

```powershell
$env:MSAL_CLIENT_ID="<real-client-id>"
$env:OFFICE_SSO_RESOURCE="api://ctrl-ai.expedient.cloud/<real-client-id>"
node tools/make-production-manifest.mjs https://ctrl-ai.expedient.cloud
```

- Output file:

```text
dist/manifest.production.xml
```

- Confirm `dist/manifest.production.xml` does not contain the placeholder app id `00000000-0000-0000-0000-000000000000` before upload.
- Confirm the manifest `WebApplicationInfo` resource matches the Entra app registration.

## Admin Center steps

1. Open Microsoft 365 Admin Center.
2. Go to `Settings` -> `Integrated apps`.
3. Choose `Upload custom apps`.
4. Select `Office Add-in`.
5. Upload `dist/manifest.production.xml`.
6. Assign users or groups for pilot rollout.
7. Finish deployment.
8. Ask pilot users to restart Excel, Word, and PowerPoint.
9. Confirm `CTRL AI` appears on the Home ribbon.

## Pilot validation

For each Office host:

- Open a normal Excel, Word, or PowerPoint file.
- Click `CTRL AI` from the ribbon.
- Confirm the task pane loads from the hosted domain.
- Configure BYOK provider settings.
- Test connection.
- Confirm Office SSO/M365 file context works for a signed-in pilot user, or that the app returns clear setup guidance if SSO is intentionally disabled.
- Read current context.
- Insert the assistant response back into the document.

## Rollout recommendation

Start with a small pilot group before org-wide deployment:

1. IT/admin pilot: validate manifest trust and domain access.
2. Power-user pilot: validate BYOK provider settings and Office context actions.
3. Department rollout: assign groups in Microsoft 365 Admin Center.
4. Broad rollout: expand assignment after monitoring support tickets and provider logs.

## Updating the add-in

For UI/server code changes:

1. Rebuild and redeploy the hosted app.
2. If manifest URLs/icons/permissions changed, regenerate and re-upload `dist/manifest.production.xml`.
3. If only web code changed and manifest did not, users should get the new task pane on reload/reopen.

For ribbon icon changes, bump the manifest version and use new icon filenames because Office caches command assets aggressively.
