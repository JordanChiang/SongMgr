using System;
using System.Collections.Generic;
using System.Linq;

namespace CrazyKTV_SongMgr
{
    /// <summary>
    /// 套用錢櫃編號: works out how the 歌庫's songs move to their 錢櫃 numbers before anything is written, then writes
    /// it in one transaction. A song sitting on a number another song gets moves to a fresh 歌曲編號; moves are ordered
    /// so each lands on a number that is free at that point, with a stop on a spare number to break a swap.
    /// </summary>
    static class CashboxIdPlan
    {
        public sealed class Song
        {
            public Song(string id, string lang, string singer, string name)
            {
                Id = id; Lang = lang; Singer = singer; Name = name;
            }

            public string Id { get; }
            public string Lang { get; }
            public string Singer { get; }
            public string Name { get; }
            public override string ToString() => Lang + "-" + Singer + "-" + Name;
        }

        public enum Reason
        {
            /// <summary>The song takes its 錢櫃 number.</summary>
            Cashbox,
            /// <summary>The song was on a number another song takes, and gets a new 歌曲編號.</summary>
            MadeRoom,
            /// <summary>A stop on a spare number, so two songs can swap.</summary>
            Spare,
        }

        public sealed class Move
        {
            public Move(Song song, string from, string to, Reason reason)
            {
                Song = song; From = from; To = to; Why = reason;
            }

            public Song Song { get; }
            public string From { get; }
            public string To { get; }
            public Reason Why { get; }
        }

        public sealed class Result
        {
            public List<Move> Moves { get; } = new List<Move>();
            /// <summary>Matches left out because another song already has that 錢櫃 number, or the song already has one.</summary>
            public List<(Song Song, string CashboxId)> Skipped { get; } = new List<(Song, string)>();
        }

        /// <param name="library">Every song in the 歌庫.</param>
        /// <param name="matches">A song of <paramref name="library"/> and the 錢櫃 number it should have.</param>
        /// <param name="allocatorFor">The 歌曲編號配置 to use, given the numbers that must not be handed out.</param>
        public static Result Plan(IReadOnlyList<Song> library, IEnumerable<(Song Song, string CashboxId)> matches, Func<IEnumerable<string>, SongIdAllocator> allocatorFor)
        {
            var result = new Result();

            // Who gets which 錢櫃 number: one number per song and one song per number, lowest number first.
            var target = new Dictionary<Song, string>();
            var owner = new Dictionary<string, Song>();
            foreach (var m in matches.OrderBy(m => m.CashboxId, StringComparer.Ordinal).ThenBy(m => m.Song.Id, StringComparer.Ordinal))
            {
                if (target.ContainsKey(m.Song) || owner.ContainsKey(m.CashboxId))
                {
                    if (!(target.TryGetValue(m.Song, out string t) && t == m.CashboxId)) result.Skipped.Add(m);
                    continue;
                }
                target[m.Song] = m.CashboxId;
                owner[m.CashboxId] = m.Song;
            }

            // A song on a number another song takes, and with none of its own, gets a number outside every 錢櫃 number.
            SongIdAllocator allocator = allocatorFor(library.Select(s => s.Id).Concat(owner.Keys));
            var final = new Dictionary<Song, (string To, Reason Why)>();
            foreach (var t in target)
            {
                if (t.Key.Id != t.Value) final[t.Key] = (t.Value, Reason.Cashbox);
            }
            foreach (Song s in library)
            {
                if (!target.ContainsKey(s) && owner.TryGetValue(s.Id, out Song taker) && taker != s)
                {
                    final[s] = (allocator.Take(s.Lang), Reason.MadeRoom);
                }
            }

            // Write order: a move goes once nothing sits on its number any more; a swap goes through a spare number.
            var at = library.GroupBy(s => s.Id).ToDictionary(g => g.Key, g => g.Count());
            var current = library.ToDictionary(s => s, s => s.Id);
            var pending = final.Keys.OrderBy(s => s.Id, StringComparer.Ordinal).ToList();
            while (pending.Count > 0)
            {
                Song next = pending.FirstOrDefault(s => !at.ContainsKey(final[s].To));
                if (next != null)
                {
                    Step(result, at, current, next, final[next].To, final[next].Why);
                    pending.Remove(next);
                }
                else
                {
                    Song first = pending[0];
                    Step(result, at, current, first, allocator.Take(first.Lang), Reason.Spare);
                }
            }
            return result;
        }

        static void Step(Result result, Dictionary<string, int> at, Dictionary<Song, string> current, Song song, string to, Reason why)
        {
            string from = current[song];
            result.Moves.Add(new Move(song, from, to, why));
            if (--at[from] == 0) at.Remove(from);
            at[to] = at.TryGetValue(to, out int n) ? n + 1 : 1;
            current[song] = to;
        }

        /// <summary>
        /// Writes <paramref name="moves"/> in order, the 我的最愛 moving with each song, all or nothing.
        /// Returns false, with nothing changed, when a song was not where the plan expected it or a write failed.
        /// </summary>
        public static bool Apply(SongLibraryStore store, IEnumerable<Move> moves, out Exception error)
        {
            error = null;
            store.BeginTransaction();
            try
            {
                foreach (Move m in moves)
                {
                    if (store.RenumberSong(m.From, m.To, m.Song.Name, m.Song.Lang) != 1)
                    {
                        throw new InvalidOperationException("歌庫裡找不到 " + m.From + " (" + m.Song + ")");
                    }
                }
                store.Commit();
                return true;
            }
            catch (Exception ex)
            {
                store.Rollback();
                error = ex;
                return false;
            }
        }
    }
}
