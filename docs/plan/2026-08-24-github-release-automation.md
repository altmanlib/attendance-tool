# GitHub Release Automation Implementation Plan

**Goal:** Automatically create a GitHub Release and attach the Windows distribution ZIP whenever a `v*` version tag is pushed

**Architecture:** The existing Windows build workflow will retain its manual-dispatch behavior and build artifact. After it creates the ZIP, a tag-only release step will use the repository-scoped GitHub token to create a release for the triggering tag and upload that same ZIP. Branch-based manual builds never create releases

**Tech Stack:** GitHub Actions, GitHub CLI, PyInstaller, PowerShell

## Global Constraints

- Create releases only for pushed tags matching `v*`
- Keep manual workflow dispatch limited to Actions artifacts
- Attach `dist/考勤统计工具-windows.zip` to the GitHub Release
- Use the GitHub CLI preinstalled on the Windows GitHub-hosted runner
- Grant only the required `contents: write` repository permission
- Do not push commits or tags without explicit permission

---

### Task 1: Publish tagged builds as GitHub Releases

**Files:**

- Modify: `.github/workflows/build-windows.yml`
- Test: GitHub Actions run triggered by a new `v*` tag

**Interfaces:**

- Consumes: `github.ref`, `github.ref_name`, and `dist/考勤统计工具-windows.zip`
- Produces: a GitHub Release named after the pushed tag with ZIP asset `考勤统计工具-windows.zip`

- [ ] **Step 1: Permit release creation**

Replace the workflow permission block with:

```yaml
permissions:
  contents: write
```

- [ ] **Step 2: Add the idempotent tag-only release upload step after the artifact upload**

```yaml
      - name: Create or update GitHub Release
        if: startsWith(github.ref, 'refs/tags/')
        shell: pwsh
        env:
          GH_TOKEN: ${{ github.token }}
        run: |
          $asset = "dist/考勤统计工具-windows.zip"
          gh release view $env:GITHUB_REF_NAME *> $null
          if ($LASTEXITCODE -eq 0) {
            gh release upload $env:GITHUB_REF_NAME $asset --clobber
          } else {
            gh release create $env:GITHUB_REF_NAME $asset --title $env:GITHUB_REF_NAME --generate-notes
          }
```

- [ ] **Step 3: Validate workflow security and syntax**

Run: `uvx --from zizmor zizmor .github/workflows/build-windows.yml`

Expected: `No findings to report`

- [ ] **Step 4: Commit**

```bash
git add .github/workflows/build-windows.yml
git commit -m "ci: publish tagged Windows builds as releases"
```

### Task 2: Document release behavior

**Files:**

- Modify: `README.md` in the GitHub Actions section
- Test: Markdown review and matching tag pattern inspection

**Interfaces:**

- Consumes: version tag pattern `v*`
- Produces: release workflow instructions for internal operators

- [ ] **Step 1: Document the release trigger and delivery location**

Add this paragraph:

```markdown
推送形如 `v1.0.0` 的 Git 标签会自动创建同名 GitHub Release，并将 `考勤统计工具-windows.zip` 附加到 Release。普通用户可在仓库的「Releases」页面直接下载该安装包
```

- [ ] **Step 2: Verify documented behavior matches workflow conditions**

Run: `rg -n "v1.0.0|refs/tags|考勤统计工具-windows.zip" README.md .github/workflows/build-windows.yml`

Expected: the README names the tag format and ZIP, and the workflow contains the tag condition and ZIP path

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "docs: describe automated release publishing"
```

## Self-Review

- Spec coverage: Task 1 creates a release with the packaged ZIP for every pushed version tag. Task 2 tells users where to download it
- Placeholder scan: no placeholders or deferred implementation items are present
- Type consistency: the ZIP path is `dist/考勤统计工具-windows.zip` in both workflow and plan, and the artifact filename is `考勤统计工具-windows.zip`
