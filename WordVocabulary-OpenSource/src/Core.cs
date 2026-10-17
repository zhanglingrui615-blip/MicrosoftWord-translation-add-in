using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace WordVocabulary
{
    public class Entry
    {
        public string Word { get; set; }
        public string Sentence { get; set; }
        public string Note { get; set; }
        public string Created { get; set; }
    }

    public static class Storage
    {
        public static readonly string Folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordVocabulary");
        public static JavaScriptSerializer Json = new JavaScriptSerializer();
        public static string PathFor(string name) { Directory.CreateDirectory(Folder); return Path.Combine(Folder, name); }
        public static void AtomicWrite(string path, string data)
        {
            string temp = path + ".tmp";
            File.WriteAllText(temp, data, new UTF8Encoding(false));
            if (File.Exists(path)) File.Replace(temp, path, path + ".bak");
            else File.Move(temp, path);
        }
        public static List<Entry> Load()
        {
            string path = PathFor("vocabulary.json");
            return File.Exists(path) ? Json.Deserialize<List<Entry>>(File.ReadAllText(path)) : new List<Entry>();
        }
        public static void Save(List<Entry> entries) { AtomicWrite(PathFor("vocabulary.json"), Json.Serialize(entries)); }
        public static Entry FindEntry(IEnumerable<Entry> entries, string word)
        {
            string normalized = ReadingInteraction.NormalizeWord((word ?? "").Trim());
            return entries.FirstOrDefault(e => String.Equals(ReadingInteraction.NormalizeWord((e.Word ?? "").Trim()), normalized, StringComparison.OrdinalIgnoreCase));
        }
        public static void Add(string word, string sentence, string note)
        {
            var entries = Load();
            if (FindEntry(entries, word) != null) return;
            entries.Add(new Entry { Word = word, Sentence = sentence, Note = note, Created = DateTimeOffset.Now.ToString("o") });
            Save(entries);
        }
        public static void Remove(string word)
        {
            var entries = Load();
            bool changed = false;
            Entry found;
            while ((found = FindEntry(entries, word)) != null) { entries.Remove(found); changed = true; }
            if (changed) Save(entries);
        }
        public static void SaveKeys(string key, string secret)
        {
            var bytes = Encoding.UTF8.GetBytes(Json.Serialize(new string[] { key, secret }));
            AtomicWrite(PathFor("credentials.dat"), Convert.ToBase64String(ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser)));
        }
        public static string[] LoadKeys()
        {
            string path = PathFor("credentials.dat");
            if (!File.Exists(path)) return new string[] { "", "" };
            return Json.Deserialize<string[]>(Encoding.UTF8.GetString(ProtectedData.Unprotect(Convert.FromBase64String(File.ReadAllText(path)), null, DataProtectionScope.CurrentUser)));
        }
        public static string Csv(string text)
        {
            text = text ?? "";
            if (text.Length > 0 && "=+-@\t\r\n".IndexOf(text[0]) >= 0) text = "'" + text;
            return "\"" + text.Replace("\"", "\"\"") + "\"";
        }
    }

    public class Definition
    {
        public string Meaning;
        public string Phonetic;
    }

    public static class Youdao
    {
        private static readonly HttpClient Client = new HttpClient { Timeout = TimeSpan.FromSeconds(15) };
        private static string ReadPhonetic(Dictionary<string, object> fields, params string[] names)
        {
            foreach (string name in names)
            {
                object value;
                if (fields.TryGetValue(name, out value))
                {
                    string text = Convert.ToString(value).Trim().Trim('/', '[', ']');
                    if (!String.IsNullOrWhiteSpace(text)) return "/" + text + "/";
                }
            }
            return "";
        }
        public static string Phonetics(Dictionary<string, object> fields)
        {
            string uk = ReadPhonetic(fields, "uk-phonetic", "ukPhonetic");
            string us = ReadPhonetic(fields, "us-phonetic", "usPhonetic");
            if (uk.Length > 0 || us.Length > 0) return String.Join("；", new[] { uk.Length > 0 ? "英 " + uk : "", us.Length > 0 ? "美 " + us : "" }.Where(s => s.Length > 0));
            return ReadPhonetic(fields, "phonetic");
        }
        public static string Input(string q) { return q.Length <= 20 ? q : q.Substring(0, 10) + q.Length + q.Substring(q.Length - 10); }
        public static string Sign(string key, string q, string salt, string time, string secret)
        {
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(key + Input(q) + salt + time + secret))).Replace("-", "").ToLowerInvariant();
        }
        public static Definition Parse(string json)
        {
            var root = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(json);
            string code = root.ContainsKey("errorCode") ? Convert.ToString(root["errorCode"]) : "unknown";
            if (code != "0") throw new InvalidOperationException("有道返回错误 " + code + "。请检查应用绑定、额度和密钥。");
            var result = new Definition { Meaning = "暂无释义", Phonetic = "" };
            if (root.ContainsKey("translation")) result.Meaning = String.Join("\n", ((System.Collections.IEnumerable)root["translation"]).Cast<object>().Select(Convert.ToString));
            // Some accounts return optional dictionary fields. Translation is always the fallback.
            if (root.ContainsKey("basic"))
            {
                var basic = root["basic"] as Dictionary<string, object>;
                if (basic != null && basic.ContainsKey("explains")) result.Meaning = String.Join("\n", ((System.Collections.IEnumerable)basic["explains"]).Cast<object>().Select(Convert.ToString));
                if (basic != null) result.Phonetic = Phonetics(basic);
            }
            if (String.IsNullOrWhiteSpace(result.Phonetic)) result.Phonetic = Phonetics(root);
            return result;
        }
        public static async Task<Definition> Lookup(string word, CancellationToken token, bool retryCompound = true)
        {
            word = ReadingInteraction.IsWord(word) ? ReadingInteraction.NormalizeWord(word) : word;
            string[] keys = Storage.LoadKeys();
            if (String.IsNullOrWhiteSpace(keys[0]) || String.IsNullOrWhiteSpace(keys[1])) throw new InvalidOperationException("请先在“翻译助手 → 接口设置”填写有道应用 ID 和密钥。");
            string salt = Guid.NewGuid().ToString("N");
            string time = ((long)(DateTime.UtcNow - new DateTime(1970, 1, 1)).TotalSeconds).ToString(System.Globalization.CultureInfo.InvariantCulture);
            var fields = new Dictionary<string, string> { { "q", word }, { "from", "en" }, { "to", "zh-CHS" }, { "appKey", keys[0] }, { "salt", salt }, { "curtime", time }, { "signType", "v3" }, { "sign", Sign(keys[0], word, salt, time, keys[1]) } };
            using (var content = new FormUrlEncodedContent(fields))
            using (var response = await Client.PostAsync("https://openapi.youdao.com/api", content, token))
            {
                response.EnsureSuccessStatusCode();
                var result = Parse(await response.Content.ReadAsStringAsync());
                // Some translation services echo an unrecognised hyphenated compound.
                if (retryCompound && ReadingInteraction.IsWord(word) && word.Contains("-") && !System.Text.RegularExpressions.Regex.IsMatch(result.Meaning, @"[\u3400-\u9fff]"))
                {
                    var phrase = await Lookup(word.Replace('-', ' '), token, false);
                    if (!String.IsNullOrWhiteSpace(phrase.Meaning)) result.Meaning = phrase.Meaning;
                    if (String.IsNullOrWhiteSpace(result.Phonetic)) result.Phonetic = phrase.Phonetic;
                }
                return result;
            }
        }
    }
}
