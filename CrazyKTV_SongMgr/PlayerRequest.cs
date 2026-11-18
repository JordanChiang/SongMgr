using System;

namespace CrazyKTV_SongMgr
{
    /// <summary>What the player window needs to open a song, and where its changes go.</summary>
    class PlayerRequest
    {
        /// <summary>Uses Id, Lang, Singer, SongName, Track, Volume, ReplayGain.</summary>
        public SongRecord Song { get; set; }
        public string FilePath { get; set; }

        /// <summary>
        /// Takes the 伴唱 track the user saved in the player. Called once the player has closed and released the file,
        /// since saving a track can rename the file. Null disables 更新音軌.
        /// </summary>
        public Action<string> TrackSaved { get; set; }

        /// <summary>Takes the volume the user saved in the player, right away. Null disables 更新音量.</summary>
        public Action<string> VolumeSaved { get; set; }
    }
}
