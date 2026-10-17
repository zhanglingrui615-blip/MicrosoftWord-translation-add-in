using System;
using System.Linq;

namespace WordVocabulary
{
    public static class WordMarks
    {
        public static void ApplyRange(object rawRange, bool saved)
        {
            try
            {
                dynamic range = rawRange;
                int style = (int)range.Font.Underline;
                if (saved) range.Font.Underline = 11;
                else if (style == 11) range.Font.Underline = 0;
                else if (style == 9999999)
                    for (int i = 1; i <= (int)range.Characters.Count; i++)
                        if ((int)range.Characters[i].Font.Underline == 11) range.Characters[i].Font.Underline = 0;
            }
            catch { } // Protected/read-only content must not interrupt the toggle.
        }
        public static void SyncWord(object rawApplication, string word, bool saved)
        {
            if (String.IsNullOrWhiteSpace(word) || word.Length > 255 || word.Contains('\r') || word.Contains('\n')) return;
            string normalized = ReadingInteraction.NormalizeWord(word);
            string[] variants = new[] { word, normalized, normalized.Replace('-', '\u2010'), normalized.Replace('-', '\u2011') }.Distinct().ToArray();
            try
            {
                dynamic app = rawApplication;
                for (int document = 1; document <= (int)app.Documents.Count; document++)
                {
                    try
                    {
                        foreach (dynamic firstStory in app.Documents[document].StoryRanges)
                        {
                            dynamic story = firstStory;
                            while (story != null)
                            {
                                foreach (string variant in variants) SyncStory((object)story, variant, saved);
                                story = story.NextStoryRange;
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }
        private static void SyncStory(object rawStory, string word, bool saved)
        {
            try
            {
                dynamic story = rawStory;
                dynamic remaining = story.Duplicate;
                int end = (int)story.End;
                int cursor = (int)story.Start;
                while (cursor < end)
                {
                    remaining.SetRange(cursor, end);
                    dynamic find = remaining.Find;
                    find.ClearFormatting();
                    bool found = find.Execute(FindText: word, MatchCase: false, MatchWholeWord: true, MatchWildcards: false, MatchSoundsLike: false, MatchAllWordForms: false, Forward: true, Wrap: 0, Format: false);
                    if (!found) break;
                    int next = (int)remaining.End;
                    ApplyRange((object)remaining, saved);
                    if (next <= cursor) break;
                    cursor = next;
                }
            }
            catch { }
        }
    }
}
