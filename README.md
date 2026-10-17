Word拾词插件 · Word Shici Add-in


## English

Look up English words while reading in Microsoft Word, listen to their pronunciation, and save vocabulary for later review.

The add-in uses a local dictionary by default and requires no API credentials for offline lookup. You can optionally connect your own Youdao translation API to translate missing words or selected sentences. The current interface and dictionary definitions are in Chinese; this bilingual README does not imply an English interface.

### Features

- **Offline dictionary:** approximately 767,000 entries with Chinese definitions derived from ECDICT. Displays the original word, available phonetic transcription, and dictionary meanings.
- **Click to look up:** enable translation mode, click an English word to select it and display a nearby dictionary card, and click blank space to dismiss it.
- **Pronunciation:** words are read aloud when their cards open, with manual playback also available through installed Windows English voices.
- **Vocabulary notebook:** add notes, save a word, and click the saved-state button again to remove it.
- **Vocabulary marks:** saved words receive wavy underlines; removing a word clears the corresponding marks.
- **Spreadsheet export:** export to Excel (.xlsx) or CSV, including the word, original sentence, personal note, and date added.
- **Optional online translation:** use your own Youdao credentials for words missing from the dictionary, phrases, and selected sentences.

> Some dictionary entries have no phonetic transcription. The add-in displays available data without inventing pronunciation. Offline definitions do not automatically determine the meaning in context.

### Compatibility

| Environment | Status |
| --- | --- |
| Windows desktop Word / Office LTSC 2024 | Tested |
| Word 2016, 2019, 2021, and Microsoft 365 on Windows | Not yet tested; feedback welcome |
| 32-bit / 64-bit Office | Installer includes both registration paths; other combinations require testing |
| macOS, Word for the web, and WPS | Currently unsupported |

Installation uses the Windows .NET Framework compiler. Visual Studio, Python, and a separate dictionary service are not required. Pronunciation requires an installed Windows English voice.

### Download and installation

1. On the repository homepage, select **Code → Download ZIP**, then extract the archive.
2. Find the project directory containing `Update.cmd`, `install.ps1`, `src`, and `data`. Keep the entire directory together; do not copy only the scripts or DLL.
3. **Save your documents and close all Word windows.**
4. Double-click `Update.cmd` and wait for compilation and installation to finish.
5. Reopen Word, find the **翻译助手** (Translation Assistant) tab, and enable **翻译模式** (Translation Mode).

The installer registers the add-in for the current Windows user and normally does not require administrator privileges. Keep the project directory in place after installation. If you move it, close Word and run `Update.cmd` again from the new location.

Alternatively, open a terminal in the project directory and run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
```

### Usage

With **翻译模式** enabled:

1. Click an English word to view its definitions and hear its pronunciation.
2. Optionally enter a note and click **加入生词本** (Add to Vocabulary).
3. Click **已加入生词本** (Added to Vocabulary) again to remove the word.
4. Click blank space in the document to dismiss the card and clear the word selection.
5. Open **生词本** (Vocabulary Notebook) to review, delete, or export entries.

Drag to select a phrase or sentence to query that selection. In default offline mode, sentences only display a configuration message. Requests to Youdao occur only after online translation is enabled. Each selection supports up to 6,000 characters.

Vocabulary marking changes underline formatting in open documents and persists when you save them. Removing entries clears the wavy underline style used by the add-in while preserving ordinary underlines. Closed documents are not opened or modified in the background. Existing underlines that use the same wavy style may also be cleared.

### Optional: bring your own Youdao API

**Online translation is disabled by default. No author credentials or shared translation allowance are provided.**

To enable it:

1. Obtain your own AppKey and AppSecret for the Youdao text translation API.
2. Enter them under **翻译助手 → 接口设置** (Translation Assistant → API Settings).
3. Check **开启在线翻译（使用自己的有道额度）** (Enable online translation using your own Youdao allowance) and save.

| Query | Default offline mode | Online translation enabled |
| --- | --- | --- |
| Word found locally | Local result; no API request | Still local; no API request |
| Word missing locally | Configuration message | Your own Youdao API |
| Phrase or sentence | Configuration message | Your own Youdao API |

Uncheck the option and save to return to offline mode. Online usage is charged to the user's own Youdao account according to the provider's rules. Offline dictionary lookup uses no API allowance.

This version integrates the **Youdao text translation API**, which does not guarantee phonetic transcription or multiple dictionary meanings. There is no general-purpose language model API configuration in this version.

### Data and privacy

- Vocabulary, encrypted credentials, and the online setting are stored under `%LOCALAPPDATA%\WordVocabulary` for the current user.
- Credentials are protected with Windows current-user encryption and are not stored in the source directory.
- Online translation sends only the actively queried text, not the entire document. Local dictionary hits do not trigger translation requests.
- This repository contains no author's credentials, vocabulary records, or personal exports. Do not distribute your user data directory.

This version shares the add-in identity and user data directory with the earlier personal version. Installation replaces the registered add-in path; do not install both versions simultaneously. Online translation starts disabled on first use, then retains the user's setting.

### Update and uninstall

**Update:** save and close Word, place the complete new version in its intended location, and run `Update.cmd`. Vocabulary and credentials remain in the separate user data directory.

**Uninstall:** save and close Word, open a terminal in the project directory, and run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -Uninstall
```

Uninstalling removes add-in registration but retains local user data.

### Troubleshooting

**`install.ps1` cannot be found:** open the terminal in the directory containing the script, or double-click `Update.cmd`.

**The installer asks you to close Word:** save and exit all Word windows. If necessary, check Task Manager for a remaining `WINWORD.EXE` process. End it only after confirming your documents are saved.

**The 翻译助手 tab is missing:** restart Word, then go to **File → Options → Add-ins → Manage: COM Add-ins → Go** and check whether `Word Vocabulary` is loaded. If it reports a runtime error, include your Office version, bitness, and an error screenshot in your report.

**A word has no phonetic transcription or cannot be found:** the upstream dictionary does not cover every pronunciation, new word, inflected form, or compound. You may use your own Youdao API for missing entries, but that text translation API also does not guarantee pronunciation data.

**There is no sound:** check the system volume and installed Windows English voices. Dictionary lookup still works without speech.

### Development

Built with C# / .NET Framework and loaded into Word through COM.

```text
src/                         Add-in source code
data/lexicon.json            Embedded common-word dictionary
data/full/                   Full dictionary buckets and checksums
tests/                       Checks and dictionary maintenance scripts
build.ps1                    Build script
install.ps1                  Current-user install / uninstall script
Update.cmd                   Installation entry point
THIRD-PARTY-NOTICES.txt       Third-party data licenses
```

Build and run checks:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\test.ps1
```

Checks cover offline/online routing, vocabulary toggling, and Excel export without calling a real paid API. UI behavior on other Office versions still requires testing in Word.

### Feedback and license

Report problems through repository **Issues**, including reproduction steps, Windows/Word version, Office bitness, and relevant screenshots. Do not submit API credentials or documents containing private information.

The add-in code is available under the [MIT License](LICENSE). Dictionary data is derived from [ECDICT](https://github.com/skywind3000/ECDICT) and other sources; see [THIRD-PARTY-NOTICES.txt](THIRD-PARTY-NOTICES.txt) for copyright and license details. Keep these notices when redistributing.

---

## 中文说明

在 Word 中阅读英文时，点击单词即可查词、听发音，并把需要复习的单词加入生词本。

默认使用本地词库，无需配置 API。需要翻译未收录的词或句子时，可以选择接入自己的有道翻译 API。

## 功能

- **离线查词**：包含 ECDICT 衍生的约 76.7 万个中文释义词条，显示原词、词库已有音标及多义项。
- **点击即查**：开启翻译模式后，点击英文单词自动选中，在附近显示词条；点击空白关闭词条。
- **单词朗读**：弹出单词卡片时自动朗读，也可手动播放，使用 Windows 本机英语语音。
- **生词本**：支持添加笔记，点击“加入生词本”保存，再次点击“已加入生词本”即可移出。
- **生词标记**：加入的词使用波浪下划线标记，移出时清除相应标记。
- **表格导出**：支持 Excel（.xlsx）和 CSV，包含单词、原句、个人笔记、收藏时间。
- **可选在线翻译**：使用者配置自己的有道 API 后，可翻译本地未收录的词、短语和选中的句子。

> 词库并非每个词都有音标，插件只显示已有数据，不编造读音。离线释义包含词库提供的义项，不自动判断句中义。

## 兼容性

| 环境 | 状态 |
| --- | --- |
| Windows 桌面版 Word / Office LTSC 2024 | 已实际验证 |
| Windows 版 Word 2016、2019、2021、Microsoft 365 | 尚未实测，欢迎反馈 |
| 32 位 / 64 位 Office | 安装脚本包含两种注册方式，其他组合仍需实测 |
| Mac、网页版 Word、WPS | 当前不支持 |

安装使用 Windows 的 .NET Framework 编译器，无需安装 Visual Studio、Python 或额外的词典服务。朗读需要本机已安装英语语音。

## 下载与安装

1. 在本仓库首页点击 **Code → Download ZIP**，下载后右键选择“全部解压”。
2. 找到包含 `Update.cmd`、`install.ps1`、`src` 和 `data` 的项目目录。请保留整个目录，不要只复制脚本或 DLL。
3. **先保存文档，关闭所有 Word 窗口。**
4. 双击 `Update.cmd`，等待编译与安装完成。
5. 重新打开 Word，进入 **翻译助手** 选项卡，点击 **翻译模式**。

安装按当前 Windows 用户注册加载项，通常无需管理员权限。安装后不要随意移动或删除项目目录；若移动，请关闭 Word 后从新位置重新运行 `Update.cmd`。

如果双击无法安装，也可以在项目文件夹空白处右键打开终端，执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
```

## 使用方法

开启 **翻译模式** 后：

1. 点击英文单词，查看释义并听发音。
2. 有需要时填写笔记，点击 **加入生词本**。
3. 已加入的词会显示 **已加入生词本**，再次点击即可移出。
4. 点击文本空白处，关闭词条并取消单词选中。
5. 在 **生词本** 中查看记录、删除词条或导出表格。

拖动选中短语或句子后，会查询选中内容。默认离线模式下，句子仅显示配置提示；开启在线翻译后才会请求有道，一次最多支持 6000 字符。

生词标记会修改打开文档的下划线格式，并随文档保存。移出时清除插件使用的波浪下划线样式，保留普通下划线；关闭的文档不会被后台打开修改。若文档本来使用相同的波浪下划线样式，该样式也可能被清除。

## 可选：接入自己的有道 API

**默认关闭在线翻译，不附带作者的密钥，也不提供公共翻译额度。**

需要在线功能时：

1. 自行申请有道文本翻译 API 的应用 ID（AppKey）和应用密钥（AppSecret）。
2. 在 Word 的 **翻译助手 → 接口设置** 中填写自己的信息。
3. 勾选 **开启在线翻译（使用自己的有道额度）**，点击保存。

| 查询内容 | 默认离线模式 | 开启在线翻译后 |
| --- | --- | --- |
| 本地收录的单词 | 本地返回，不请求 API | 仍然本地返回，不请求 API |
| 本地未收录的单词 | 显示配置提示 | 使用自己的有道 API |
| 短语或句子 | 显示配置提示 | 使用自己的有道 API |

取消勾选并保存即可恢复离线模式。有道调用的费用或额度由使用者自己的账户承担，具体以服务商规则为准；离线词库查询没有 API 用量。

当前接入的是**有道文本翻译接口**，在线返回不保证包含音标和多个词典义项。此版本没有通用大模型接口配置入口。

## 数据与隐私

- 生词本、接口密钥和在线开关保存在当前用户的 `%LOCALAPPDATA%\WordVocabulary` 中。
- 密钥通过 Windows 当前用户加密保护，不保存在源码目录。
- 在线翻译只发送主动查询的待译内容，不发送整个文档；本地命中的词不发起翻译请求。
- 本仓库不包含作者的密钥、生词本或个人导出文件。分享项目时，请勿附带自己的用户数据目录。

本版本与此前的个人版本共用加载项标识和用户数据目录，安装后会替换加载项的注册位置，请勿同时安装两个版本。在线开关首次默认关闭，后续保留用户自己的设置。

## 更新与卸载

**更新：**保存并关闭 Word，将新版完整文件放好后运行 `Update.cmd`。本机生词本和接口设置保存在独立用户目录中。

**卸载：**保存并关闭 Word，在项目目录打开终端，执行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -Uninstall
```

卸载仅移除加载项注册，保留本机用户数据。

## 常见问题

**提示 `install.ps1` 不存在？**  
当前终端没有位于项目目录。请从包含该脚本的文件夹打开终端，或直接双击 `Update.cmd`。

**提示先关闭 Word？**  
保存并退出所有 Word 窗口后再安装。如果仍提示，检查任务管理器中是否有残留 `WINWORD.EXE`；确认文档已保存再结束进程。

**Word 中没有“翻译助手”选项卡？**  
先重新打开 Word，再到“文件 → 选项 → 加载项”，在底部选择“COM 加载项 → 转到”，查看 `Word Vocabulary` 是否已加载。若提示运行错误，请反馈 Office 版本、位数和错误截图。

**为什么某些词没有音标，或者查不到？**  
上游词库并非完整覆盖所有读音、新词、词形和复合词。未收录词可选择使用自己的有道 API，但文本翻译接口同样不保证返回音标。

**为什么没有声音？**  
请检查系统音量及 Windows 是否安装了英语语音。查词功能不依赖朗读成功。

## 开发

项目使用 C# / .NET Framework，通过 COM 加载到 Word。

```text
src/                         插件源码
data/lexicon.json            常用词库（随程序内嵌）
data/full/                   完整词库分桶与校验清单
tests/                       检查程序与词库维护脚本
build.ps1                    编译脚本
install.ps1                  当前用户安装 / 卸载脚本
Update.cmd                   安装入口
THIRD-PARTY-NOTICES.txt       第三方数据许可
```

编译与检查：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\test.ps1
```

检查涵盖离线查询与在线回退路由、生词本切换和 Excel 导出，不调用真实付费 API。不同 Office 版本的界面和交互仍需实机验证。

## 反馈与许可

遇到问题请在仓库 **Issues** 中描述复现步骤，并附 Windows / Word 版本、Office 位数和必要截图。请勿提交 API 密钥或带有私人内容的文档。

插件代码采用 [MIT License](LICENSE)。词库来源于 [ECDICT](https://github.com/skywind3000/ECDICT) 等项目，版权和许可见 [THIRD-PARTY-NOTICES.txt](THIRD-PARTY-NOTICES.txt)。再分发时请保留许可文件。
