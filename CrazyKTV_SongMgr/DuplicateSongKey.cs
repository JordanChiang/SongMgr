using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace CrazyKTV_SongMgr
{
    /// <summary>
    /// 重複歌曲: two songs are the same song when their keys under one <see cref="Rule"/> are equal. Every rule compares
    /// each field the same way:
    /// <list type="bullet">
    /// <item>語系 exactly;</item>
    /// <item>歌名 trimmed, whatever the case (and, under the bracket rules, with anything in brackets dropped);</item>
    /// <item>歌曲類別 whatever the case, an empty 類別 matching only an empty one;</item>
    /// <item>歌手 as the set of its singers: split as a 合唱 whatever its 歌手類別 (special singer names stay whole),
    /// each trimmed, whatever the case, and taken as its 歌手群組; the order of a 合唱 does not matter.</item>
    /// </list>
    /// The rules differ only in which fields they compare.
    /// </summary>
    static class DuplicateSongKey
    {
        public sealed class Rule
        {
            private Rule(bool singer, bool songType, bool ignoreBrackets)
            {
                Singer = singer; SongType = songType; IgnoreBrackets = ignoreBrackets;
            }

            public bool Singer { get; }
            public bool SongType { get; }
            public bool IgnoreBrackets { get; }

            /// <summary>重複歌曲: 語系, 歌名, 歌手 and 歌曲類別. Also what 加歌 and 歌庫查詢 editing count as a duplicate.</summary>
            public static readonly Rule Full = new Rule(true, true, false);
            /// <summary>重複歌曲 (忽略歌曲類別).</summary>
            public static readonly Rule IgnoreSongType = new Rule(true, false, false);
            /// <summary>重複歌曲 (忽略括號及歌曲類別).</summary>
            public static readonly Rule IgnoreBracketsAndSongType = new Rule(true, false, true);
            /// <summary>重複歌曲 (忽略歌手及歌曲類別) and 重複歌曲 (合唱歌曲): 語系 and 歌名 only.</summary>
            public static readonly Rule NameOnly = new Rule(false, false, false);
            /// <summary>同名歌曲 (忽略括號及歌曲類別): 語系 and 歌名 with its brackets dropped.</summary>
            public static readonly Rule NameOnlyIgnoreBrackets = new Rule(false, false, true);
        }

        private static readonly Regex Brackets = new Regex(@"[\{\(\[｛（［【].+?[】］）｝\]\)\}]", RegexOptions.Compiled);

        private const char Separator = '\u0001';

        public static string Of(Rule rule, string lang, string singer, string songName, string songType, SongSearchReference reference)
        {
            string name = songName ?? "";
            if (rule.IgnoreBrackets) name = Brackets.Replace(name, "");
            name = name.Trim().ToLowerInvariant();

            var key = new List<string> { lang ?? "", name };
            if (rule.Singer) key.Add(SingerKey(singer ?? "", reference));
            if (rule.SongType) key.Add((songType ?? "").ToLowerInvariant());
            return string.Join(Separator.ToString(), key);
        }

        /// <summary>The singers of a song as one text: each 歌手群組 once, in a fixed order.</summary>
        public static string SingerKey(string singer, SongSearchReference reference)
        {
            return string.Join("&", reference.SplitChorus(singer)
                                             .Select(s => reference.GroupOf(s))
                                             .Where(s => s != "")
                                             .Distinct(StringComparer.Ordinal)
                                             .OrderBy(s => s, StringComparer.Ordinal));
        }
    }
}
