using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.Text;

namespace PSEmailer
{
    public class GhostScript
    {
        // Set variables to be used in the class
        private readonly ArrayList _gsParams = new ArrayList();
        private IntPtr[] _gsArgPtrs;
        private GCHandle _gsArgPtrsHandle;
        private GCHandle[] _gsArgStrHandles;
        private IntPtr _gsInstancePtr;

        // ReSharper disable InconsistentNaming
        [DllImport("gsdll64.dll")]
        private static extern int gsapi_new_instance(out IntPtr pinstance, IntPtr caller_handle);
        // ReSharper restore InconsistentNaming

        [DllImport("gsdll64.dll")]
        private static extern int gsapi_init_with_args(IntPtr instance, int argc, IntPtr argv);

        [DllImport("gsdll64.dll")]
        private static extern int gsapi_exit(IntPtr instance);

        [DllImport("gsdll64.dll")]
        private static extern void gsapi_delete_instance(IntPtr instance);

        public void AddParam(string param)
        {
            _gsParams.Add(param);
        }

        public void Execute()
        {
            try
            {
                // Create GS Instance (GS-API)
                gsapi_new_instance(out _gsInstancePtr, IntPtr.Zero);

                // Build Argument Arrays
                _gsArgStrHandles = new GCHandle[_gsParams.Count];
                _gsArgPtrs = new IntPtr[_gsParams.Count];

                // Populate Argument Arrays
                for (int i = 0; i < _gsParams.Count; i++)
                {
                    _gsArgStrHandles[i] = GCHandle.Alloc(Encoding.ASCII.GetBytes(_gsParams[i].ToString()), GCHandleType.Pinned);
                    _gsArgPtrs[i] = _gsArgStrHandles[i].AddrOfPinnedObject();
                }

                // Allocate memory that is protected from Garbage Collection
                _gsArgPtrsHandle = GCHandle.Alloc(_gsArgPtrs, GCHandleType.Pinned);

                // Init args with GS instance (GS-API)
                gsapi_init_with_args(_gsInstancePtr, _gsArgStrHandles.Length, _gsArgPtrsHandle.AddrOfPinnedObject());
            }
            catch (Exception)
            {
                throw new ExternalException("GhostScript failed.");
            }
            finally
            {
                try
                {
                    // Free unmanaged memory
                    for (int i = 0; i < _gsArgStrHandles.Length; i++)
                    {
                        _gsArgStrHandles[i].Free();
                    }
                    _gsArgPtrsHandle.Free();

                    // Exit the api (GS-API)
                    gsapi_exit(_gsInstancePtr);

                    // Delete GS Instance (GS-API)
                    gsapi_delete_instance(_gsInstancePtr);
                }
                catch (Exception)
                {
                    throw new ExternalException("GhostScript ceanup failed.");
                }
            }
        }
    }
}