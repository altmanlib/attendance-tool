# Windows GitHub Actions Build Implementation Plan

**Goal:** Add a GitHub Actions workflow that packages the desktop application as a Windows executable and exposes the complete distributable folder as a downloadable ZIP artifact

**Architecture:** A single Windows-hosted workflow will run only when manually dispatched or when a version tag is pushed. It installs the locked uv environment, invokes PyInstaller in `--onedir` mode, compresses the generated application directory, and uploads the archive as a workflow artifact. The README will direct internal users to the Actions artifact

**Tech Stack:** GitHub Actions, `windows-latest`, uv, PyInstaller, PowerShell

## Global Constraints

- Build Windows executables on a Windows GitHub-hosted runner
- Use the dependency versions locked in `uv.lock`
- Package with PyInstaller `--windowed --onedir`
- Deliver the complete application directory as a ZIP archive
- Do not push commits or create remote repositories without explicit permission

---

### Task 1: Add the Windows packaging workflow

**Files:**

- Create: `.github/workflows/build-windows.yml`
- Test: GitHub Actions workflow dispatch after the commit is pushed to GitHub

**Interfaces:**

- Consumes: `pyproject.toml`, `uv.lock`, and `main.py`
- Produces: an artifact named `attendance-tool-windows` containing `考勤统计工具-windows.zip`

- [ ] **Step 1: Create the workflow trigger and Windows job**

```yaml
name: Build Windows executable

on:
  workflow_dispatch:
  push:
    tags:
      - "v*"

jobs:
  build:
    name: Package Windows application
    runs-on: windows-latest
```

- [ ] **Step 2: Check out source code and install locked dependencies**

```yaml
    steps:
      - name: Check out repository
        uses: actions/checkout@11d5960a326750d5838078e36cf38b85af677262 # v4.4.0
        with:
          persist-credentials: false

      - name: Install uv
        uses: astral-sh/setup-uv@d0cc045d04ccac9d8b7881df0226f9e82c39688e # v6.8.0
        with:
          enable-cache: false

      - name: Install dependencies
        run: uv sync --all-groups --frozen
```

- [ ] **Step 3: Build and archive the distributable application**

```yaml
      - name: Build executable
        run: >-
          uv run pyinstaller --noconfirm --clean --windowed --onedir
          --name "考勤统计工具" main.py

      - name: Create distribution archive
        shell: pwsh
        run: Compress-Archive -Path "dist/考勤统计工具" -DestinationPath "dist/考勤统计工具-windows.zip"
```

- [ ] **Step 4: Upload the archive as a downloadable artifact**

```yaml
      - name: Upload Windows package
        uses: actions/upload-artifact@ea165f8d65b6e75b540449e92b4886f43607fa02 # v4.6.2
        with:
          name: attendance-tool-windows
          path: dist/考勤统计工具-windows.zip
          if-no-files-found: error
          retention-days: 30
```

- [ ] **Step 5: Validate workflow syntax locally**

Run: `uvx --from zizmor zizmor .github/workflows/build-windows.yml`

Expected: the command reports no workflow security findings

- [ ] **Step 6: Commit**

```bash
git add .github/workflows/build-windows.yml
git commit -m "ci: build Windows executable in GitHub Actions"
```

### Task 2: Document artifact retrieval

**Files:**

- Modify: `README.md` after the "打包为 Windows `.exe`" section
- Test: inspect rendered Markdown in GitHub after push

**Interfaces:**

- Consumes: workflow artifact `attendance-tool-windows`
- Produces: operator instructions for manually starting a build and downloading the ZIP

- [ ] **Step 1: Add the GitHub Actions build instructions**

```markdown
## 使用 GitHub Actions 打包

提交代码到 GitHub 后，打开仓库的「Actions」→「Build Windows executable」→「Run workflow」启动构建。构建完成后，在该运行记录底部的「Artifacts」下载 `attendance-tool-windows`，解压后将完整的 `考勤统计工具` 文件夹交付给用户
```

- [ ] **Step 2: Verify the documented artifact name matches the workflow**

Run: `rg -n "attendance-tool-windows" README.md .github/workflows/build-windows.yml`

Expected: one reference in the README and one `name:` field in the workflow

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "docs: explain GitHub Actions Windows builds"
```

## Self-Review

- Spec coverage: Task 1 creates a Windows-only GitHub Actions packaging workflow and stores the complete onedir distribution in an artifact. Task 2 tells internal operators how to trigger and retrieve it
- Placeholder scan: no placeholders or deferred implementation items are present
- Type consistency: the artifact name `attendance-tool-windows` is identical in the workflow and documentation
