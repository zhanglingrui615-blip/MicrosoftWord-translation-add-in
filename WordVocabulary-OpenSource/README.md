# 翻译助手：Word 离线查词插件

Windows 桌面 Word 插件，默认离线查词，不附带任何 API 密钥。包含 ECDICT 衍生的约 76.7 万个中文释义词条，保留词库提供的音标和多义项；原词库缺少音标的词不编造读音。

## 使用

1. 下载完整项目并解压，保留 src、data 等目录，不要只下载一个脚本。
2. 保存文档并关闭所有 Word 窗口。
3. 双击 Update.cmd，脚本使用 Windows 自带的 .NET Framework 编译器编译并注册当前用户的 COM 加载项，无需管理员权限或 Visual Studio。
4. 打开 Word，在“翻译助手”选项卡开启“翻译模式”。点击单词查看词条；点击空白关闭词条。生词可加入或再次点击移出，支持 Excel / CSV 导出。

适用于 Windows 桌面版 Word；开发时使用 Office LTSC 2024。其他 Office 版本需自行验证；不支持 Mac、网页版 Word 或 WPS。朗读使用 Windows 已安装的英语语音。

## 可选在线翻译：使用自己的 API

默认关闭在线翻译。本地命中的词始终离线返回；缺词和句子在默认模式下只显示说明，不发送翻译请求。

需要在线功能时，在“接口设置”填写自己的有道文本翻译 AppKey、AppSecret，勾选“开启在线翻译”后保存。只有本地未命中的词、短语和句子才会发送给有道；费用或额度由使用者自己的有道账户承担。取消勾选并保存即可恢复离线模式。这里集成的是有道文本翻译接口，不保证在线结果含音标或多个义项，也不是任意模型通用接口。

项目不提供公共密钥或代付服务。离线词库查询没有 API 用量；有道额度按其服务规则计量，并不等同于大模型 token。

## 数据与隐私

生词本、加密密钥、在线开关保存在当前用户的 %LOCALAPPDATA%\WordVocabulary 下，不在项目文件夹。密钥使用 Windows 当前用户加密。在线请求仅提交主动查询的待译内容，不上传整个文档。开源包不包含作者密钥、生词本或查询记录。

此版本与原版共用加载项标识和用户数据目录；安装会替换注册位置，请勿同时安装两个版本。首次使用开源版的在线开关默认关闭；若此前已使用过开源版，则保留自己的开关状态。卸载保留本机用户数据。

加入生词本会修改打开文档中对应词语的波浪下划线；移出后清除该样式，普通下划线保留。词库格式本身无法提供上下文词义判断，需阅读多个义项选择合适含义。

## 开发与检查

运行 powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 编译。
运行 powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\test.ps1 执行不调用付费 API 的路由与生词本/导出检查。
卸载：关闭 Word 后运行 powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -Uninstall。

src 为 C# 源码；data/full 为按词首分桶的完整离线词库；manifest.json 记录词库信息与校验值。tests 中的 Python 维护脚本用于重新下载/生成词库，安装和日常使用无需 Python。

上传 GitHub 时上传本目录中的文件与目录；不要上传本机用户数据目录、bin、密钥文件或个人导出表格。ZIP 是分发压缩包，如通过 GitHub 网页提交源码，应先解压再上传其中内容。

## 许可

插件代码见 LICENSE。词库来源、版权及 MIT 许可见 THIRD-PARTY-NOTICES.txt；分发词库时请一并保留。
