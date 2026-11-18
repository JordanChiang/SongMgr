using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
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

        /// <summary>
        /// Finds a song's duplicates, under <see cref="Rule.Full"/>, in a list of "語系|歌手|歌名|歌曲類別" texts, as 加歌,
        /// 歌曲分析 and 歌庫查詢 editing keep the 歌庫's songs. The list may grow between lookups (歌曲分析 adds each song
        /// it keeps); entries already in it never change. One index per list, shared by every thread.
        /// </summary>
        public sealed class Index
        {
            private static readonly ConditionalWeakTable<List<string>, Index> indexes = new ConditionalWeakTable<List<string>, Index>();

            private readonly List<string> source;
            private readonly SongSearchReference reference;
            private readonly Dictionary<string, List<int>> byKey = new Dictionary<string, List<int>>(StringComparer.Ordinal);
            private int seen;

            public Index(List<string> songData, SongSearchReference reference)
            {
                source = songData;
                this.reference = reference;
            }

            /// <summary>The index of <paramref name="songData"/>, matching against SongMgr's current 歌手群組 and special singer names.</summary>
            public static Index For(List<string> songData)
            {
                lock (indexes)
                {
                    if (!indexes.TryGetValue(songData, out Index index))
                    {
                        index = new Index(songData, SongSearchReference.Capture(SongSearchKind.Duplicate));
                        indexes.Add(songData, index);
                    }
                    return index;
                }
            }

            /// <summary>The positions in the list of the songs that duplicate this one, first one first.</summary>
            public List<int> Find(string lang, string singer, string songName, string songType)
            {
                string key = Of(Rule.Full, lang, singer, songName, songType, reference);
                lock (byKey)
                {
                    lock (source)
                    {
                        if (seen > source.Count) { byKey.Clear(); seen = 0; }
                        for (; seen < source.Count; seen++) Add(source[seen], seen);
                    }
                    return byKey.TryGetValue(key, out List<int> found) ? new List<int>(found) : new List<int>();
                }
            }

            private void Add(string songData, int position)
            {
                string[] parts = songData.Split('|');
                string key = (parts.Length < 4)
                    ? songData
                    : Of(Rule.Full, parts[0], string.Join("|", parts.Skip(1).Take(parts.Length - 3)), parts[parts.Length - 2], parts[parts.Length - 1], reference);

                if (!byKey.TryGetValue(key, out List<int> positions))
                {
                    positions = new List<int>();
                    byKey.Add(key, positions);
                }
                positions.Add(position);
            }
        }
    }
}
