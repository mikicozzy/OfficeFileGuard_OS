using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;

namespace CrocoIconOverlayShellExtension
{
    [ComVisible(false)]
    [ComImport]
    [Guid("0C6C4200-C589-11D0-999A-00C04FD655E1")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IShellIconOverlayIdentifier
    {
        [PreserveSig]
        int IsMemberOf([MarshalAs(UnmanagedType.LPWStr)] string path,
                       [MarshalAs(UnmanagedType.U4)] int attributes);

        [PreserveSig]
        int GetOverlayInfo(IntPtr iconFileBuffer, int iconFileBufferSize,
                           out int iconIndex, out uint flags);

        [PreserveSig]
        int GetPriority(out int priority);
    }

    [ComVisible(true)]
    [Guid("098AA80F-2F55-4AA8-B95F-8AF77120E479")]
    [ClassInterface(ClassInterfaceType.None)]
    public class CrocoOfficeOverlayHandler : IShellIconOverlayIdentifier
    {
        // --- CAMPI STATICI ---

        // Queue condivisa con tutte le istanze (static perché Explorer può creare più istanze della classe)
        private static readonly ConcurrentQueue<string> _verificationQueue = new ConcurrentQueue<string>();

        // generazione di un cancellation token ('_cts') utilizzabile per fermare il worker
        private static readonly CancellationTokenSource _cts = new CancellationTokenSource();

        // generazione del task del worker
        private static Task _workerTask;

        // semaforo per evitare istanze parallele di explorer con esecuzione parallela di questo worker
        private static readonly object _workerLock = new object();

        // Path dell'icona .ico da usare come iconoverlay  (
        private static readonly string IconPath = Path.Combine(
            Path.GetDirectoryName(typeof(CrocoOfficeOverlayHandler).Assembly.Location),
            "lucchetto_tr.ico"
        );

        // --- P/INVOKE ---
        // la funzione SHChangeNotify di shell32.dll notifica al sistema le modifiche al file system, consentendo ad Explorer di aggiornare le icone quando necessario.
        [DllImport("shell32.dll")]
        private static extern void SHChangeNotify(int wEventId, uint uFlags,
                                                   IntPtr dwItem1, IntPtr dwItem2);

        // --- CONSTRUCTOR ---
        public CrocoOfficeOverlayHandler()
        {
            StartWorkerIfNeeded();
        }


        // --- INTERFACE IShellIconOverlayIdentifier ---
        public int IsMemberOf(string path, int attributes)
        {
            try
            {
                // to speedup: only supported extension are verified
                string ext = Path.GetExtension(path).ToLowerInvariant();
                if (ext != ".docx" && ext != ".xlsx" && ext != ".pptx")
                    return 1; // S_FALSE — non mostrare, non accodare

                bool isReservedInRegistry = ReservedFileRegistry.IsOverlayNeeded(path);
                // adding path to the queuefor asynchronous verification
                _verificationQueue.Enqueue(path);
                return isReservedInRegistry ? 0 : 1; // S_OK = 0 (show), S_FALSE = 1 (don't show)
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"isMemberOf EXCEPTION: {ex.Message}");
                return 1; // S_FALSE = non mostrare
            }
        }

        public int GetOverlayInfo(IntPtr iconFileBuffer, int iconFileBufferSize,
                                  out int iconIndex, out uint flags)
        {
            try
            {
                // Copia il path dell'icona nel buffer fornito da Explorer
                var iconPathChars = IconPath.ToCharArray();
                int charsToCopy = Math.Min(iconPathChars.Length, iconFileBufferSize - 1);
                for (int i = 0; i < charsToCopy; i++)
                    Marshal.WriteInt16(iconFileBuffer, i * 2, iconPathChars[i]);
                Marshal.WriteInt16(iconFileBuffer, charsToCopy * 2, 0); // null terminator

                iconIndex = 0;
                flags = 1; // ISIOI_ICONFILE
                return 0;  // S_OK
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetOverlayInfo EXCEPTION: {ex.Message}");
                iconIndex = 0;
                flags = 0;
                return -1; // E_FAIL
            }
        }

        public int GetPriority(out int priority)
        {
            priority = 0; // 0 = massima priorità
            return 0;     // S_OK
        }


        // --- ASYNC WORKER ---
        private static void StartWorkerIfNeeded()
        {
            lock (_workerLock)  // acquire the semaphore to start the worker
            {
                if (_workerTask != null && !_workerTask.IsCompleted)
                    return;  // worker already running

                _workerTask = Task.Run(() =>    //strats a new thread for the worker task
                {
                    while (!_cts.Token.IsCancellationRequested)  // loop until cancellation is requested
                    {
                        if (_verificationQueue.TryDequeue(out string path))  // tryDequeue returns the next item of the queue (false if the queue is empty)
                        {
                            try
                            {
                                if (!File.Exists(path)) //if the file has been deleted, it is removed from the registry and Explorer is notified to update the icons
                                {
                                    if (ReservedFileRegistry.IsReserved(path))  // the file is present in the registry
                                    {
                                        CrocoOfficeFileGuard.Core.ReservedFileManager.UnregisterReservedFile(path);   // remove the file from the registry
                                        NotifyExplorer(path);  // notify Explorer to update the icons for this path
                                    }
                                    continue;    // go to the start of the loop to get the next path to verify
                                }

                                bool isReservedInFile = CrocoOfficeFileGuard.Core.ReservedFileManager.IsReservedInFile(path);  // check the file to see if it contains the reserved property
                                bool isReservedInRegistry = ReservedFileRegistry.IsReserved(path);      // check the registry to see if the file is marked as reserved

                                if (isReservedInFile != isReservedInRegistry)    // there is a mismatch (to be solved)
                                {
                                    if (isReservedInFile)   // if the file is actually reserved (then is not so in the registry)
                                        CrocoOfficeFileGuard.Core.ReservedFileManager.RegisterReservedFile(path);   // add the file to the registry
                                    else
                                        CrocoOfficeFileGuard.Core.ReservedFileManager.UnregisterReservedFile(path); // remove the file from the registry

                                    NotifyExplorer(path);    // notify Explorer to update the icons for this path
                                }
                            }

                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"StartWorkerIfNeeded EXCEPTION: {ex.Message}");
                            }
                        }
                        else    // queue is empty, wait a bit before checking again
                        {
                            Thread.Sleep(200);  // if no paths to verify, wait a bit before checking again
                        }
                    }
                }, _cts.Token);
            }
        }


        private static void NotifyExplorer(string path)
        {
            IntPtr pathPtr = Marshal.StringToHGlobalUni(path);
            try
            {
                // SHCNE_UPDATEITEM = 0x00002000, SHCNF_PATHW = 0x0005
                SHChangeNotify(0x00002000, 0x0005, pathPtr, IntPtr.Zero);
            }
            finally
            {
                Marshal.FreeHGlobal(pathPtr);
            }
        }
    }
}