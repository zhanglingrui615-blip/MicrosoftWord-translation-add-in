using System;
using System.Drawing;
using System.Text.RegularExpressions;

namespace WordVocabulary
{
    public static class ReadingInteraction
    {
        public const int MaxSelectionLength = 6000;
        public const string WordPattern = @"[A-Za-z]+(?:['’\-\u2010\u2011][A-Za-z]+)*";
        public static bool IsWord(string text) { return Regex.IsMatch(text, "^" + WordPattern + "$"); }
        public static string NormalizeWord(string text) { return text.Replace('\u2010', '-').Replace('\u2011', '-').Replace('’', '\''); }
        public static Match WordNearCaret(string text, int caret)
        {
            foreach (Match match in Regex.Matches(text, WordPattern))
                if (caret >= match.Index && caret <= match.Index + match.Length) return match;
            return Match.Empty;
        }
        public static string VisiblePhonetic(string text)
        {
            if (String.IsNullOrWhiteSpace(text) || text.Contains("http")) return "";
            var matches = Regex.Matches(text, @"(?:(?:英|美)\s*)?/[^/\r\n]+/");
            var parts = new System.Collections.Generic.List<string>();
            foreach (Match match in matches) parts.Add(match.Value);
            return String.Join("；", parts);
        }
        public static bool IsGlyphHit(Point point, Rectangle bounds, string text)
        {
            return !String.IsNullOrWhiteSpace(text) && bounds.Width > 0 && bounds.Height > 0 && bounds.Contains(point);
        }
        public static Point PlaceCard(Rectangle anchor, Size card, Rectangle screen)
        {
            int x = Math.Max(screen.Left, Math.Min(anchor.Left, screen.Right - card.Width));
            int y = anchor.Bottom + 5;
            if (y + card.Height > screen.Bottom) y = anchor.Top - card.Height - 5;
            return new Point(x, Math.Max(screen.Top, Math.Min(y, screen.Bottom - card.Height)));
        }
    }
}
