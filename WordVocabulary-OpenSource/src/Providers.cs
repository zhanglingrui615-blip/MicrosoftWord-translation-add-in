using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
namespace WordVocabulary
{
    public static class OnlineTranslation
    {
        public static bool Enabled
        {
            get { try { return File.ReadAllText(Storage.PathFor("online-enabled.txt")).Trim() == "true"; } catch { return false; } }
            set { Storage.AtomicWrite(Storage.PathFor("online-enabled.txt"), value ? "true" : "false"); }
        }
    }
    public static class DictionaryProvider
    {
        public static Task<Definition> Lookup(string word, string sentence, CancellationToken token, IProgress<Definition> progress = null)
        {
            return LookupWithFallback(word, token, OnlineTranslation.Enabled, (query, cancel) => Youdao.Lookup(query, cancel));
        }
        public static async Task<Definition> LookupWithFallback(string word, CancellationToken token, bool onlineEnabled, Func<string, CancellationToken, Task<Definition>> fallback)
        {
            token.ThrowIfCancellationRequested();
            Definition local;
            if (ReadingInteraction.IsWord(word) && LocalLexicon.TryLookup(word, out local)) return local;
            if (!onlineEnabled) return new Definition { Meaning = ReadingInteraction.IsWord(word)
                ? "本地词库暂未收录。需要在线查词时，请在接口设置中填写自己的有道密钥并开启在线翻译。"
                : "离线模式仅提供单词查词。需要句子翻译时，请在接口设置中填写自己的有道密钥并开启在线翻译。", Phonetic = "" };
            return await fallback(word, token);
        }
    }
}
