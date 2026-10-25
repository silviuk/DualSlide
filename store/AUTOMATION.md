# Microsoft Store automation

The **Publish to Microsoft Store** GitHub Action follows Luminark's release and Partner Center API approach. It builds the exact selected release commit, verifies the MSIX identity and screenshot hashes, uploads the package and listing images, commits the API submission, and reports its status. Certification remains Microsoft's decision; only `Published` means the app is live.

## Run it

1. Open GitHub Actions → **Publish to Microsoft Store** → **Run workflow** on `main`.
2. Enter an exact published stable tag, for example `v0.5.4`.
3. Leave **submit** off to build and validate without contacting Partner Center. Download the `DualSlide-store-candidate` artifact to inspect the package, screenshots, notes, and recorded source SHA.
4. Enable **submit** to upload and commit to certification. Submission is allowed only from `main`; it never runs on a pull request or an ordinary code push.

Screenshot metadata in `store/Screenshots/screenshots.json` must identify the selected version. Review or refresh the screenshots and their SHA-256 hashes before publishing a new version. The app's own tagged `RELEASE_NOTES.md` supplies What's New. The workflow checkout supplies the packaging recipe, privacy document, and listing; it builds the executable and user guide from the selected immutable source commit. This also supports the existing `v0.5.4` release, whose original tag predates the MSIX recipe.

Normal future version-tag builds also attach the MSIX and checksum to the GitHub release. They do not automatically submit to the Store.

## Publisher credentials

Configure these **Actions secrets in silviuk/DualSlide** using the existing Partner Center-associated Entra application:

| Secret | Meaning |
| --- | --- |
| `MS_STORE_TENANT_ID` | Entra tenant ID |
| `MS_STORE_CLIENT_ID` | Partner Center-associated Entra application client ID |
| `MS_STORE_CLIENT_SECRET` | Its client secret |

The MSIX API does not require the EXE/MSI API's seller-ID header. The reserved product ID `9NP8S6HC2252` and MSIX identity `SilviuVlasceanu.DualSlide` are fixed in the scripts to prevent submitting to Luminark or Lunifier by mistake. GitHub cannot reveal another repository's saved secret values through its API; sharing or configuring credentials must be authorized separately. No token, client secret, or upload SAS URL is stored in artifacts or printed in errors.

## One-time Partner Center setup

Microsoft requires a completed initial submission with age ratings before API-created updates can be based on its published submission. DualSlide's initial web draft must therefore be completed once. Set the pricing/markets, Desktop-only device availability, IARC ratings, publisher/support details, privacy URL, and required restricted-capability explanations. The Store product is already reserved.

- Product: https://partner.microsoft.com/en-us/dashboard/products/9NP8S6HC2252/overview
- Privacy URL: https://github.com/silviuk/DualSlide/blob/main/PRIVACY.md
- Support: https://github.com/silviuk/DualSlide/issues

The automation preserves these baseline settings. The older listing API ignores privacy-policy, support-contact and website URLs, so the script does not pretend to update those fields. Complete them in Partner Center's Properties page during initial setup.

API-created drafts must not subsequently be edited in Partner Center. If a run fails before commit, the summary records its draft ID. Supply that exact value as **resume_submission_id** only for that API-created draft. Other editable drafts are preserved and block submission. A committed submission is checked for the exact version and screenshots rather than recommitted or overwritten. No draft is automatically deleted.

## Local verification

```powershell
./tools/Test-StorePublishing.ps1
./build-msix.ps1
```

The tests mock the HTTP boundary and exercise failed uploads, first-submission prerequisites, explicit resume, idempotent reruns, wrong products/versions, path traversal, unsafe upload endpoints, and modified screenshots. They never contact Microsoft or need real credentials. MSIX packaging needs the Windows SDK's MakeAppx.exe; the script discovers it on GitHub's Windows runners.

Reference implementation: https://github.com/silviuk/Luminark/blob/main/.github/workflows/publish-store-and-winget.yml

Microsoft API prerequisites: https://learn.microsoft.com/en-us/windows/uwp/monetize/create-and-manage-submissions-using-windows-store-services

Microsoft submission lifecycle: https://learn.microsoft.com/en-us/windows/uwp/monetize/manage-app-submissions
