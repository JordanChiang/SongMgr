using System;
using System.IO;
using CrazyKTV_SongMgr;
using SongMgr.Tests.Support;
using Xunit;

namespace SongMgr.Tests.Media
{
    public class MediaProbeTests : IDisposable
    {
        readonly string dir = TestEnv.NewTempDir("probe");

        public void Dispose() => TestEnv.DeleteDir(dir);

        [Fact]
        public void Counts_audio_streams()
        {
            var f = MediaFixtures.Create(dir);
            Assert.Equal(1, MediaProbe.AudioStreamCount(f.one));
            Assert.Equal(2, MediaProbe.AudioStreamCount(f.two));
            Assert.Equal(3, MediaProbe.AudioStreamCount(f.three));
            Assert.Equal(0, MediaProbe.AudioStreamCount(f.empty));
        }

        [Fact]
        public void Verify_file_accepts_a_clean_file_and_rejects_a_broken_one()
        {
            var f = MediaFixtures.Create(dir);
            Assert.True(MediaProbe.VerifyFile(f.two));
            Assert.False(MediaProbe.VerifyFile(f.empty));
        }

        [Theory]
        [InlineData(50, 0.0, 50)]
        [InlineData(50, 12.0, 100)]
        [InlineData(30, -20.0, 10)]
        public void Library_range_pulls_volume_into_10_to_100(int baseVolume, double gain, int expected)
        {
            Assert.Equal(expected.ToString(), MediaProbe.SongVolume(baseVolume, gain, VolumeRange.Library));
        }

        [Fact]
        public void Normalize_range_keeps_quiet_songs_down_to_5()
        {
            Assert.Equal("5", MediaProbe.SongVolume(30, -20.0, VolumeRange.Normalize));
        }
    }
}
