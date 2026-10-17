using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Web.Script.Serialization;

namespace WordVocabulary
{
    public static class LocalLexicon
    {
        private static readonly Lazy<Dictionary<string, string[]>> Data = new Lazy<Dictionary<string, string[]>>(Load);
        private static Dictionary<string, string[]> Load()
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("WordVocabulary.Lexicon.json"))
            {
                if (stream == null) return new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
                using (var reader = new StreamReader(stream))
                {
                    var serializer = new JavaScriptSerializer { MaxJsonLength = 32 * 1024 * 1024 };
                    return new Dictionary<string, string[]>(serializer.Deserialize<Dictionary<string, string[]>>(reader.ReadToEnd()), StringComparer.OrdinalIgnoreCase);
                }
            }
        }
        public static int Count { get { return Math.Max(FullLexicon.Count, Data.Value.Count); } }
        public static bool TryLookup(string word, out Definition result)
        {
            result = null;
            word = ReadingInteraction.NormalizeWord(word);
            if (String.Equals(word, "campus-only", StringComparison.OrdinalIgnoreCase))
            {
                result = new Definition { Phonetic = "", Meaning = "adj. 仅限校园的；仅供校内使用的；仅对校内开放的。\n\n由 campus（校园）与 only（仅、只）组成，具体修饰对象取决于原句。" };
                return true;
            }
            string[] values;
            if (!Data.Value.TryGetValue(word, out values) && !FullLexicon.TryLookup(word, out values)) return false;
            result = new Definition
            {
                Phonetic = values[0],
                Meaning = values[1] + "\n\n来源：ECDICT 本地词库"
            };
            return true;
        }
    }
}
