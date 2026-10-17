using System;
using System.Threading;
using System.Threading.Tasks;
using WordVocabulary;
class PublicRoutingTests
{
    static int Main()
    {
        int calls = 0;
        Func<string,CancellationToken,Task<Definition>> fake = (q,t) => { calls++; return Task.FromResult(new Definition { Meaning = "fake" }); };
        foreach (bool enabled in new[] { false, true })
        {
            var local = DictionaryProvider.LookupWithFallback("hello", CancellationToken.None, enabled, fake).GetAwaiter().GetResult();
            if (calls != 0 || String.IsNullOrWhiteSpace(local.Meaning)) throw new Exception("Local lookup must stay offline.");
        }
        foreach (string q in new[] { "zzzznotaword999999", "This is a full sentence." })
        {
            var offline = DictionaryProvider.LookupWithFallback(q, CancellationToken.None, false, fake).GetAwaiter().GetResult();
            if (calls != 0 || !offline.Meaning.Contains("接口设置")) throw new Exception("Disabled online fallback called.");
            var online = DictionaryProvider.LookupWithFallback(q, CancellationToken.None, true, fake).GetAwaiter().GetResult();
            if (online.Meaning != "fake") throw new Exception("Enabled fallback failed.");
            calls = 0;
        }
        var cts = new CancellationTokenSource(); cts.Cancel();
        try { DictionaryProvider.LookupWithFallback("hello", cts.Token, true, fake).GetAwaiter().GetResult(); return 1; }
        catch (OperationCanceledException) { }
        Console.WriteLine("PASS: local lookup, offline missing words/sentences, explicit online fallback and cancellation. No real API called.");
        return 0;
    }
}
