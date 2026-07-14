# Comments for OneNote

Embedded, movable comments for the Windows desktop version of Microsoft OneNote.

Comments for OneNote is a community fork of [Steven M. Cohn's OneMore](https://github.com/stevencohn/OneMore). It keeps the full OneMore feature set and adds comments that are attached to selected text, shown inside an embedded right-side pane, and reconnected when the original text moves.

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

## Install a release

1. Install the official [OneMore 7.2.0](https://github.com/stevencohn/OneMore/releases) if needed.
2. Download `Comments-for-OneNote-win-x64-v*.zip` from this repository's Releases page.
3. Extract the complete ZIP and close OneNote.
4. Open PowerShell in the extracted folder and run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
```

The installer uses a per-user COM registration override. It does not overwrite the official OneMore installation in `Program Files`.

To remove this community build and return to official OneMore:

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

> **安装前提：请先安装官方 OneMore 7.2.0 x64，再安装本项目。**

Comments for OneNote 是基于 [OneMore](https://github.com/stevencohn/OneMore) 的社区分支，为 Windows 桌面版 OneNote 增加内嵌评论。评论保存在页面隐藏元数据中，原文移动后会根据原文及上下文重新连接。

使用方法：在同一段落内选中文字，点击 **OneMore → 添加评论**，或右键选择 **添加评论**。点击 **评论面板** 可显示或隐藏右侧内嵌面板。

首个发布版本支持 64 位 OneNote。卸载脚本会移除社区版本的用户级覆盖并恢复官方 OneMore。
