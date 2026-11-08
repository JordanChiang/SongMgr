using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using CrazyKTV_MediaKit.DirectShow.MediaPlayers;
using DirectShowLib;

namespace CrazyKTV_SongMgr
{
    /// <summary>
    /// Finds out whether the VMR9 video renderer works here by building one the way MediaKit does for playback.
    /// Querying Direct3D9 or matching adapter names is not enough: with the adapter disabled, the HAL device resolves
    /// to the Microsoft Basic Render Driver and CreateDevice succeeds, and only the VMR9's SetD3DDevice then fails.
    /// </summary>
    internal static class VideoEnvCheck
    {
        // MediaUriPlayer creates its renderer with two streams.
        private const int StreamCount = 2;

        private static readonly IntPtr AllocatorUserId = new IntPtr(1);

        public static bool IsVmr9Available()
        {
            try
            {
                return TryCreateVmr9();
            }
            catch (FileNotFoundException)
            {
                // Codec_x86 is not installed, which says nothing about the display environment.
                return true;
            }
            catch
            {
                return false;
            }
        }

        // Not inlined so that a missing MediaKit dll surfaces inside IsVmr9Available's try block.
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static bool TryCreateVmr9()
        {
            IBaseFilter filter = (IBaseFilter)new VideoMixingRenderer9();
            Vmr9Allocator allocator = null;
            try
            {
                IVMRFilterConfig9 config = (IVMRFilterConfig9)filter;
                if (config.SetNumberOfStreams(StreamCount) < 0) return false;
                if (config.SetRenderingMode(VMR9Mode.Renderless) < 0) return false;

                IVMRSurfaceAllocatorNotify9 notify = (IVMRSurfaceAllocatorNotify9)filter;
                allocator = new Vmr9Allocator();
                if (notify.AdviseSurfaceAllocator(AllocatorUserId, allocator) < 0) return false;
                return allocator.AdviseNotify(notify) >= 0;
            }
            finally
            {
                // Once AdviseNotify has run, Dispose also releases the filter, so the second release may find it gone.
                // Cleanup must not change the result.
                try
                {
                    allocator?.Dispose();
                    Marshal.ReleaseComObject(filter);
                }
                catch
                {
                }
            }
        }
    }
}
