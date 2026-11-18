using CrazyKTV_SongMgr;
using Xunit;

namespace SongMgr.Tests.Media
{
    public class PlayerOutputTests
    {
        [Theory]
        [InlineData("1", true)]
        [InlineData("2", false)]
        [InlineData("", false)]
        public void Setting_1_chooses_vmr9_and_anything_else_evr(string setting, bool vmr9)
        {
            Assert.Equal(vmr9 ? VideoOutput.Vmr9 : VideoOutput.Evr, PlayerOutput.FromSetting(setting));
        }

        [Fact]
        public void Warns_when_vmr9_is_chosen_and_unavailable()
        {
            Assert.True(PlayerOutput.NeedsVmr9Warning(VideoOutput.Vmr9, () => false));
            Assert.False(PlayerOutput.NeedsVmr9Warning(VideoOutput.Vmr9, () => true));
        }

        [Fact]
        public void Evr_neither_warns_nor_probes_vmr9()
        {
            bool probed = false;
            Assert.False(PlayerOutput.NeedsVmr9Warning(VideoOutput.Evr, () => { probed = true; return false; }));
            Assert.False(probed);
        }
    }
}
