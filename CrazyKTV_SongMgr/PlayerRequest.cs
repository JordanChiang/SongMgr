namespace CrazyKTV_SongMgr
{
    /// <summary>What the player window needs to open a song.</summary>
    class PlayerRequest
    {
        /// <summary>Uses Id, Lang, Singer, SongName, Track, Volume, ReplayGain.</summary>
        public SongRecord Song { get; set; }
        public string FilePath { get; set; }

        /// <summary>Row of the calling grid, sent back with a track change.</summary>
        public string RowIndex { get; set; }

        /// <summary>Who opened the player: "SongQuery", "SongQueryEdit", "SongAdd" or "NormVol".</summary>
        public string Source { get; set; }
    }

    /// <summary>A 伴唱 track the user picked in the player, for the grid or edit panel that opened it.</summary>
    class PlayerTrackUpdate
    {
        public string Source { get; set; }
        public string RowIndex { get; set; }
        public string Track { get; set; }
    }
}
