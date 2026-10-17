using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Web.Script.Serialization;

namespace WordVocabulary
{
    // Read only the requested two-character bucket; never load the entire dictionary into Word.
    public static class FullLexicon
    {
        private static readonly string Folder = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..", "data", "full"));
        private static readonly Dictionary<string, Dictionary<string, string[]>> Cache = new Dictionary<string, Dictionary<string, string[]>>();
        private static readonly LinkedList<string> Recency = new LinkedList<string>();
        public static string BucketName(string word)
        {
            word = ReadingInteraction.NormalizeWord(word).ToLowerInvariant();
            var name = new StringBuilder();
            foreach (byte value in Encoding.UTF8.GetBytes(word.Substring(0, Math.Min(2, word.Length)))) name.Append(value.ToString("x2"));
            return name.ToString() + ".json";
        }
        public static int Count
        {
            get
            {
                string manifest = Path.Combine(Folder, "manifest.json");
                if (!File.Exists(manifest)) return 0;
                try
                {
                    var data = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(manifest));
                    return Convert.ToInt32(data["entries"]);
                }
                catch (IOException) { return 0; }
                catch (ArgumentException) { return 0; }
            }
        }
        public static bool TryLookup(string word, out string[] values)
        {
            values = null;
            word = ReadingInteraction.NormalizeWord(word).ToLowerInvariant();
            if (word.Length == 0) return false;
            string bucket = BucketName(word);
            lock (Cache)
            {
                Dictionary<string, string[]> entries;
                if (Cache.TryGetValue(bucket, out entries))
                {
                    Recency.Remove(bucket); Recency.AddLast(bucket);
                    return entries.TryGetValue(word, out values);
                }
                string path = Path.Combine(Folder, bucket);
                if (!File.Exists(path)) return false;
                try
                {
                    var serializer = new JavaScriptSerializer { MaxJsonLength = 32 * 1024 * 1024 };
                    entries = serializer.Deserialize<Dictionary<string, string[]>>(File.ReadAllText(path, Encoding.UTF8));
                    if (entries == null) return false;
                }
                catch (IOException) { return false; }
                catch (ArgumentException) { return false; }
                if (Cache.Count >= 4) { Cache.Remove(Recency.First.Value); Recency.RemoveFirst(); }
                Cache[bucket] = entries; Recency.AddLast(bucket);
                return entries.TryGetValue(word, out values);
            }
        }
    }
}
