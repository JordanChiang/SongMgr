using System.Windows.Forms;

namespace CrazyKTV_SongMgr
{
    /// <summary>播放器: plays one song in its own window while the window that opened it is hidden.</summary>
    internal static class SongPlayer
    {
        /// <summary>Shows the player for <paramref name="request"/> and hides <paramref name="owner"/> until it closes.</summary>
        public static void Open(Form owner, PlayerRequest request)
        {
            DShowForm player = new DShowForm(owner, request);
            player.Show();
            owner.Hide();
        }
    }
}
