# Maintenance

This page is for maintainers. It covers how Dependabot updates merge and release on their own,
and the one-time GitHub setup that the chain needs.

## Dependency updates

Dependabot keeps the Composer packages and the GitHub Actions up to date. A minor, patch, or
security update merges and releases with no person involved:

1. Every Monday at 05:00 Eastern, Dependabot opens one pull request per ecosystem. It groups the
   minor and patch updates (`.github/dependabot.yml`). A security update arrives as soon as a
   Dependabot alert opens.
2. `dependabot-auto-merge.yml` arms auto-merge on a minor or patch update, and on a security
   update that is not a major bump. The pull request merges when the required CI checks pass.
3. CI runs again on `main`. When it passes, `auto-release.yml` tags the next patch version as a
   GitHub Release, for example `v1.5.1` to `v1.5.2`.
4. Packagist reads the new tag through its GitHub webhook.

A person still decides these cases:

- A major update, including a major security update.
- A change to the PHP floor, `"php": ">=8.2"` in `composer.json`. `dependabot.yml` ignores it,
  because the CI matrix in `ci.yml` moves with it.
- A Release when `main` holds a commit since the last tag that is not Dependabot's.
  `auto-release.yml` does not tag it, so a person reviews and tags that work.

`auto-release.yml` also skips a Release when the updates change only `.github/`. A GitHub
Actions update changes nothing that Composer installs.

### Differences from the gcgov application pattern

This setup is a copy of the one in `gcgov/lab-api`, from the `dependabot-auto-release` skill in
`gcgov/deploy`. This repository is a Composer library, not a deployed application, so three
parts differ:

| Part | `gcgov/lab-api` | This repository |
| --- | --- | --- |
| Composer `versioning-strategy` | `increase-if-necessary`, because `composer.lock` takes minor updates | `increase`, because a library commits no `composer.lock` |
| Composer `cooldown` | none | 3 days, because each update releases with no person involved |
| End of the chain | `auto-release.yml` starts `release.yml`, which deploys through `gcgov/deploy` | `auto-release.yml` stops at the tag, and Packagist reads it |

### One-time setup

Until the settings below exist, Dependabot still opens pull requests, but nothing merges on
its own. The auto-merge workflow logs a warning that names the missing secrets.

A repository admin makes each change in the GitHub web UI. Do the steps in order.

#### 1. Turn on auto-merge and squash merges

1. Open **Settings** → **General**, and scroll to **Pull Requests**.
2. Turn on **Allow squash merging**. `auto-release.yml` identifies a Dependabot change by the
   author of the squash commit, `dependabot[bot]`.
3. Turn on **Allow auto-merge**.

#### 2. Add a ruleset that requires the CI checks

GitHub's auto-merge waits only for required checks. Without a required check, GitHub refuses
to arm auto-merge, and the auto-merge workflow fails.

1. Open **Settings** → **Rules** → **Rulesets**.
2. Click **New ruleset** → **New branch ruleset**.
3. Name it, for example `main`. Set **Enforcement status** to **Active**.
4. Leave **Bypass list** empty. A bypass actor merges at once, before CI ends.
5. Under **Target branches**, click **Add target** → **Include default branch**.
6. Turn on **Require status checks to pass**.
7. Click **Add checks**, and add each CI job name. The matrix job counts once per PHP version:
   - `PHPStan`
   - `PHPUnit (PHP 8.2)`
   - `PHPUnit (PHP 8.3)`
   - `PHPUnit (PHP 8.4)`
   - `PHPUnit (PHP 8.5)`
8. Leave **Require branches to be up to date before merging** off. With it on, a grouped pull
   request waits until Dependabot rebases it.
9. Click **Create**.

A check name appears in the search box only after CI ran once with that name. If a name is
absent, open any pull request first, and let CI run.

When you add a PHP version to the CI matrix, add its check here too.

#### 3. Turn on Dependabot alerts and security updates

1. Open **Settings** → **Advanced Security**. On some accounts, the page is **Code security**.
2. Turn on **Dependency graph**, if it is off.
3. Turn on **Dependabot alerts**.
4. Turn on **Dependabot security updates**.

#### 4. Register the merge GitHub App

The merge needs a token from a GitHub App, not `GITHUB_TOKEN`. A merge that `GITHUB_TOKEN`
makes fires no push event, so CI never runs on `main`, and `auto-release.yml` never starts.
Also, `GITHUB_TOKEN` cannot read Dependabot alerts.

This repository belongs to a personal account, so the App belongs to that account. If the
repository moves to the `gcgov` organization, register the App under the organization instead.

1. Open your account **Settings** → **Developer settings** → **GitHub Apps**.
2. Click **New GitHub App**.
3. Set **GitHub App name** to a unique name, for example `microsoftservices-automerge`.
4. Set **Homepage URL** to `https://github.com/andrewsauder/microsoftServices`.
5. Under **Webhook**, clear **Active**. The App receives no events.
6. Under **Permissions** → **Repository permissions**, set these four. Leave every other
   permission at **No access**:

   | Permission | Access | Why |
   | --- | --- | --- |
   | Contents | Read and write | To merge. |
   | Pull requests | Read and write | To arm auto-merge. |
   | Workflows | Read and write | GitHub refuses a GitHub Actions update from an App without it. |
   | Dependabot alerts | Read-only | To tell a security update from a version update. |

7. Under **Where can this GitHub App be installed?**, select **Only on this account**.
8. Click **Create GitHub App**.
9. On the App page, copy the **Client ID**. It starts with `Iv`.
10. Scroll to **Private keys**, and click **Generate a private key**. The browser downloads a
    `.pem` file.
11. Click **Install App** in the side menu, and click **Install** next to your account.
12. Select **Only select repositories**, and select `microsoftServices`. Click **Install**.

Do not add the App to any bypass list. Auto-merge waits for CI only because the ruleset holds
the App back too.

#### 5. Add the App's secrets as Dependabot secrets

A workflow that Dependabot starts reads only **Dependabot** secrets. An Actions secret with the
same name is empty there.

1. Open **Settings** → **Secrets and variables** → **Dependabot**.
2. Click **New repository secret**.
3. Set **Name** to `AUTOMERGE_APP_CLIENT_ID`. Paste the client ID into **Secret**. Click
   **Add secret**.
4. Click **New repository secret** again.
5. Set **Name** to `AUTOMERGE_APP_PRIVATE_KEY`. Paste the whole content of the `.pem` file,
   with the `BEGIN` and `END` lines. Click **Add secret**.
6. Delete the downloaded `.pem` file.

The GitHub CLI can add the same two secrets:

```bash
gh secret set AUTOMERGE_APP_CLIENT_ID --app dependabot --repo andrewsauder/microsoftServices
```

```bash
gh secret set AUTOMERGE_APP_PRIVATE_KEY --app dependabot --repo andrewsauder/microsoftServices < app-private-key.pem
```

#### 6. Check that GitHub Actions can create a tag

`auto-release.yml` uses `GITHUB_TOKEN` to create the tag, and needs no secret.

1. Open **Settings** → **Rules** → **Rulesets**.
2. Confirm that no tag ruleset blocks the creation of `v*` tags. If one does, add
   **GitHub Actions** to its bypass list, or narrow its target.

#### 7. Check the Packagist webhook

1. Open the package page on packagist.org, and sign in as a maintainer.
2. Confirm that the page shows no warning that the package is not auto-updated.

If the page shows the warning, follow the Packagist instructions to connect GitHub. Without the
webhook, a new tag reaches Packagist only after a manual **Update** on the package page.

#### 8. Tag the first Release by hand

`auto-release.yml` tags only when every commit since the last tag is Dependabot's. The commit
that added this setup is a person's commit, so tag a Release after it merges:

```bash
gh release create v1.5.1 --repo andrewsauder/microsoftServices --target main --generate-notes
```

The next Dependabot Release counts from that tag.

### Verify the chain

1. Open **Insights** → **Dependency graph** → **Dependabot**. Click **Check for updates** on
   one ecosystem. Otherwise, wait for Monday at 05:00 Eastern.
2. Open the **Dependabot auto-merge** run on the new pull request. Its summary reads
   **Auto-merge armed** for a minor or patch update. It reads **Not merged on its own** for a
   major update.
3. After CI passes, the pull request merges. The commit author is `dependabot[bot]`. The merger
   is the App's bot account.
4. CI runs on `main`. Then **Auto release** runs. Its summary reads **Released vX.Y.Z**, or
   **No Release:** with the reason.
5. Packagist lists the new version.

A **No Release** result is often correct. A GitHub Actions update changes only `.github/`, and
a person's commit since the last tag holds the Release for a person.

### Troubleshoot

Find the first link of the chain that did not happen. Each workflow writes its decision to the
run summary, and to a `::notice::`, `::warning::`, or `::error::` annotation.

| Symptom | Cause | Fix |
| --- | --- | --- |
| Warning: `AUTOMERGE_APP_CLIENT_ID and AUTOMERGE_APP_PRIVATE_KEY are not both set` | The secrets are absent, or they are Actions secrets. | Add both as Dependabot secrets (step 5). |
| **Mint the merge token** fails with `Not Found` | The App is not installed on this repository, or the client ID is wrong. | Install the App (step 4). Check the client ID. |
| Error: `could not arm auto-merge` | **Allow auto-merge** or **Allow squash merging** is off, or no ruleset requires a check. | Do steps 1 and 2. |
| `alert-state` is empty on a known security update | The App lacks **Dependabot alerts: Read-only**. | Add the permission, then accept it on the installation. |
| Merge fails with `refusing to allow a GitHub App to create or update workflow` | The App lacks **Workflows: Read and write**. | Add the permission, then accept it on the installation. |
| The pull request merges before CI ends | The App is on a bypass list, or a required check matches no job name. | Empty the bypass list. Match each check to a CI job name. |
| Auto-merge armed, and the pull request never merges | A required check name matches no job, so it stays **Expected**. | Fix the check names in the ruleset. |
| No CI run on `main` after a merge | `GITHUB_TOKEN` made the merge, not the App. | Confirm that **Arm auto-merge** uses `steps.token.outputs.token`. |
| **Auto release** never runs | `auto-release.yml` is not on `main`, or CI's `name:` is not `CI`. | Merge the file. Match the name. |
| `No Release: main has commits since vX.Y.Z that are not Dependabot's` | A person's commit is on `main` since the last tag. | Correct. A person tags it. |
| **Tag the Release** fails with `403` | A tag ruleset blocks `v*` tags. | Do step 6. |

The full troubleshooting table is in `gcgov/deploy`, at
`.claude/skills/dependabot-auto-release/references/troubleshooting.md`.
