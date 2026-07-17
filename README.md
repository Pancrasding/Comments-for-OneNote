# Comments for OneNote

Add Word-like comments directly inside the Windows desktop version of OneNote.

Select text, add a comment in an embedded side pane, and keep the comment attached when the original text moves—much like comments in Microsoft Word.

Comments for OneNote is a community fork of [Steven M. Cohn's OneMore](https://github.com/stevencohn/OneMore). It keeps the full OneMore feature set while adding this Word-like commenting workflow.

> This is an independent community project. It is not an official Microsoft or OneMore release.

> **Important:** This add-on is not standalone. Install the official **OneMore 7.2.0 x64** first, then install Comments for OneNote.

## Features

- Add a comment from the OneMore ribbon or text context menu.
- View comments in a pane embedded inside the OneNote window.
- Store comments in hidden page metadata so they travel and sync with the page.
- Reconnect a comment after its source text moves to another paragraph.
- Jump to, edit, resolve, reopen, and delete comments.
- Avoid unsafe guesses when identical text occurs in ambiguous locations.

## Requirements

- Windows desktop OneNote from Microsoft Office.
- OneMore 7.2.0 installed first.
- The first community build supports 64-bit OneNote/OneMore.

## One-click install

1. Install the official [OneMore 7.2.0](https://github.com/stevencohn/OneMore/releases) if needed.
2. Download `Comments-for-OneNote-win-x64-v*.zip` from this repository's Releases page.
3. Extract the complete ZIP and close OneNote.
4. Double-click **`Install.cmd`**.

OneNote will reopen automatically when installation finishes. Keep all extracted files together while installing.

If Windows prevents the double-click installer from running, open PowerShell in the extracted folder and use:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
```

The installer uses a per-user COM registration override. It does not overwrite the official OneMore installation in `Program Files`.

To remove this community build and return to official OneMore, double-click **`Uninstall.cmd`**. The equivalent manual command is:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\uninstall.ps1
```

## Use

1. Select text within one paragraph.
2. Choose **OneMore → Add Comment** or right-click and choose **Add Comment**.
3. Enter the comment and save it.
4. Use **OneMore → Comments Pane** to show or hide the embedded pane.

Cross-paragraph selections are not yet supported.

## How anchoring works

Each comment records the quote, its OneNote object ID, and surrounding prefix/suffix context. The pane periodically checks the active page. If the object ID changes after a move, it scores candidate matches using the surrounding context before reconnecting. If matches remain ambiguous, it marks the comment unattached instead of linking to the wrong text.

## Build from source

The project targets .NET Framework 4.8 and follows the upstream OneMore build layout. For 64-bit OneNote:

```powershell
msbuild OneMore\OneMore.csproj /t:Build /p:Configuration=Debug /p:Platform=x64
```

Implementation: [`CommentFeature.cs`](OneMore/Commands/Comments/CommentFeature.cs). Tests: [`CommentStoreTests.cs`](OneMoreTests/Commands/Comments/CommentStoreTests.cs).

## Upstream, copyright, and license

This repository is derived from [OneMore](https://github.com/stevencohn/OneMore), copyright © Steven M. Cohn and OneMore contributors. The full upstream Git history is retained. See [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).

Source code is distributed under the [Mozilla Public License 2.0](LICENSE). OneMore, OneNote, Microsoft, and related names and marks belong to their respective owners.

Contributions and issue reports are welcome.

---

## 中文说明

**让 OneNote 像 Word 一样给选中的文字添加批注。**

> **安装前提：请先安装官方 OneMore 7.2.0 x64，再安装本项目。**

Comments for OneNote 是基于 [OneMore](https://github.com/stevencohn/OneMore) 的社区分支，为 Windows 桌面版 OneNote 增加类似 Microsoft Word 的内嵌批注体验。批注保存在页面隐藏元数据中，原文移动后会根据原文及上下文重新连接。

使用方法：在同一段落内选中文字，点击 **OneMore → 添加评论**，或右键选择 **添加评论**。点击 **评论面板** 可显示或隐藏右侧内嵌面板。

安装方法：下载并完整解压 Release 中的 ZIP，关闭 OneNote，然后双击 **`Install.cmd`**。无需手动输入 PowerShell 命令。要卸载并恢复官方 OneMore，双击 **`Uninstall.cmd`**。

当前版本支持 64 位 OneNote。
