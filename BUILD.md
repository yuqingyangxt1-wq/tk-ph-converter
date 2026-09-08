# 怎么把代码变成 Windows .exe

有三种路径，按你的情况选最合适的。

---

## 方式 1：双击 `build_windows.bat`（最快，你有 Windows 环境就用这个）

把整个 `tk-ph-converter/` 文件夹拷到任意一台 Windows 上，双击 `build_windows.bat`。
脚本会自动建 venv、装依赖、打 exe，最后在 `dist\` 产出可分发的文件。
**耗时 3–5 分钟，无需任何手动操作。**

> 你的 Windows 虚拟机里跑这个就行。

---

## 方式 2：GitHub Actions 云端打包（无需 Windows 机器）

`.github/workflows/build-windows.yml` 已经写好了。
云端用 Windows runner 跑 PyInstaller，产物作为 artifact 上传，你直接从网页下载。

### 用法

**一次性：把代码推到 GitHub**

```bash
cd tk-ph-converter
git init
git add .
git commit -m "init: TK PH converter v2.0"
# 在 github.com 新建一个空仓库（不要勾 README/license/.gitignore），比如 tk-ph-converter
git remote add origin git@github.com:你的用户名/tk-ph-converter.git
git branch -M main
git push -u origin main
```

**之后每次要 exe：**

- **方法 A — 网页点按钮**：进 GitHub → 你的仓库 → Actions 标签 → 左栏
  `build-windows` → 右上 `Run workflow` → 等 3–5 分钟完成 →
  在该 run 页面底部 **Artifacts** 区下载 `TK菲律宾表格转化工具-windows.zip`。
- **方法 B — 改代码后自动**：直接 `git push`，Actions 自动跑。
- **方法 C — 发正式版**：打 tag `git tag v2.0 && git push --tags`，
  Actions 会自动创建 GitHub Release 并把 exe 挂上去。

zip 解压后是 `TK菲律宾表格转化工具.exe` + `assets/` 文件夹，
双击 exe 即用，整个文件夹可发给客户。

> 免费额度：公开仓库无限；私有仓库每月 2000 分钟（这个项目每次 3–5 分钟，绰绰有余）。

---

## 方式 3：本地交叉编译（Wine）— **不推荐 arm64 Mac**

理论上 `brew install --cask wine-stable` + Windows Python + Wine 下跑 PyInstaller
可以从 Mac 出 Windows exe。但在 Apple Silicon 上：

- 需要先装 Rosetta（交互式，弹窗要确认）
- Wine 本身是 x86_64 模拟，性能差
- PyInstaller bootloader 在 Wine 下偶有兼容问题
- 全程 30–60 分钟，成功率不高

所以我**没走这条路**，直接给你云端方案（方式 2）。
如果你坚持要在 Mac 本地交叉编译，告诉我，我们再花时间走 Wine。
