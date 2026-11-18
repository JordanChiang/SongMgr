using System;
using System.Windows.Forms;

namespace CrazyKTV_SongMgr
{
    /// <summary>影片輸出: the video renderer the player uses.</summary>
    internal enum VideoOutput
    {
        Vmr9,
        Evr,
    }

    /// <summary>Which 影片輸出 the settings choose, and whether this machine can use it.</summary>
    internal static class PlayerOutput
    {
        /// <summary>MainCfgPlayerOutput: "1" is VMR9, anything else EVR.</summary>
        public static VideoOutput FromSetting(string setting) => (setting == "1") ? VideoOutput.Vmr9 : VideoOutput.Evr;

        public static VideoOutput Current => FromSetting(Global.MainCfgPlayerOutput);

        /// <summary>Only VMR9 is probed, so <paramref name="vmr9Available"/> runs only when VMR9 is chosen.</summary>
        public static bool NeedsVmr9Warning(VideoOutput output, Func<bool> vmr9Available)
        {
            return output == VideoOutput.Vmr9 && !vmr9Available();
        }

        /// <summary>Warns when VMR9 is chosen and cannot be built here. Call once the settings are loaded.</summary>
        public static void WarnIfUnavailable()
        {
            // Warn only: managing the song library needs no video renderer, just the player does.
            if (!NeedsVmr9Warning(Current, VideoEnvCheck.IsVmr9Available)) return;

            MessageBox.Show(
                "此電腦無法使用 VMR9 影片輸出，以 VMR9 播放歌曲時會失敗。\n\n" +
                "常見原因是顯示卡已停用，或尚未安裝顯示卡原廠驅動程式（裝置管理員中只顯示「Microsoft 基本顯示介面卡」）。\n\n" +
                "歌庫管理功能不受影響。請啟用顯示卡或安裝驅動程式後重新啟動程式，也可以試著到設定中將播放輸出改為 EVR。",
                "顯示環境檢查", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
