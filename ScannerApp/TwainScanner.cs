using NTwain;
using NTwain.Data;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Interop;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip;

namespace ScannerApp
{
    public partial class TwainScanner
    {
        static ManualResetEvent scanCompleted = new ManualResetEvent(false);
        static readonly object transferLock = new object();
        static System.Timers.Timer transferTimer;

        /// <summary>
        /// Scanner name from the last scan operation.
        /// </summary>
        public string ScannerName { get; private set; } = string.Empty;

        /// <summary>
        /// Scanner manufacturer from the last scan operation.
        /// </summary>
        public string ScannerManufacturer { get; private set; } = string.Empty;

        /// <summary>
        /// Scanner product family/model from the last scan operation.
        /// </summary>
        public string ScannerModel { get; private set; } = string.Empty;      

        /// <summary>
        /// NTwain-based scanning to PDF.
        /// </summary>
        /// <param name="a_sourceIndex"></param>
        /// <param name="a_fullPathPdf"></param>
        /// <param name="a_Feeder"></param>
        /// <param name="a_Duplex"></param>
        /// <param name="a_color">bw, gray, color</param>
        /// <param name="a_resolution"></param>
        /// <param name="a_PageWidth"></param>
        /// <param name="a_PageHeight"></param>
        /// <returns>Empty string for success, otherwise error message.</returns>
        public string ScanToPdf(int a_sourceIndex, string a_fullPathPdf,
                 bool a_Feeder, bool a_Duplex,
                 string a_color, string a_resolution,
                 int a_PageWidth, int a_PageHeight)
        {
            string msg = string.Empty;

            // Check TWAIN platform support
            if (!PlatformInfo.Current.IsSupported)
            {
                msg = "TWAIN platform is not supported. Drivers may be missing or incompatible.";
                Logger.Log(msg);
                return msg;
            }

            try
            {
                msg = ScanToPdf2(a_sourceIndex, a_fullPathPdf,
                    a_Feeder, a_Duplex,
                    a_color, a_resolution,
                    a_PageWidth, a_PageHeight);
            }
            catch (Exception)
            {
                //Log error and return generic message to caller
            }
            return msg;
        }

        /// <summary>
        /// Main scanning implementation using NTwain.
        /// </summary>
        /// <param name="a_sourceIndex"></param>
        /// <param name="a_fullPathPdf"></param>
        /// <param name="a_Feeder"></param>
        /// <param name="a_Duplex"></param>
        /// <param name="a_color"></param>
        /// <param name="a_resolution"></param>
        /// <param name="a_PageWidth"></param>
        /// <param name="a_PageHeight"></param>
        /// <returns></returns>
        internal string ScanToPdf2(int a_sourceIndex, string a_fullPathPdf,
                bool a_Feeder, bool a_Duplex,
                string a_color, string a_resolution,
                int a_PageWidth, int a_PageHeight)
        {
            string msg = string.Empty;
            var pages = new List<Bitmap>();     // List to accumulate pages as Bitmaps

            // debounce timeout: consider scan finished when no new pages arrive for this interval
            // increased a bit to allow the device some breathing room between pages
            const double debounceMs = 1500;
            transferTimer = new System.Timers.Timer(debounceMs) { AutoReset = false };
            transferTimer.Elapsed += (sender, args) =>
            {
                // No new transfers in debounce interval -> consider scan complete
                scanCompleted.Set();
            };

            var appId = TWIdentity.CreateFromAssembly(DataGroups.Image, Assembly.GetExecutingAssembly());
            var session = new TwainSession(appId);

            session.DataTransferred += (s, e) =>
            {
                try
                {
                    // Acquire native image stream, create a clone of the bitmap, store it
                    using (var native = e.GetNativeImageStream())
                    {
                        // Defensive: protect against empty/zero-length streams which may represent a "blank" image
                        using (var ms = new MemoryStream())
                        {
                            native.CopyTo(ms);
                            if (ms.Length == 0)
                            {
                                // It's an empty transfer (blank or driver placeholder).
                                // Restart debounce so the job does not prematurely finish.
                                lock (transferLock)
                                {
                                    transferTimer.Stop();
                                    transferTimer.Start();
                                }
                                //Log("Received empty transfer (skipped).");
                                return;
                            }

                            ms.Seek(0, SeekOrigin.Begin);
                            using (var bmpFromStream = new Bitmap(ms))
                            {
                                // Clone so the Bitmap lives independently of the stream
                                var cloned = new Bitmap(bmpFromStream);

                                lock (transferLock)
                                {
                                    pages.Add(cloned);
                                    // restart debounce timer for multi-page scans
                                    transferTimer.Stop();
                                    transferTimer.Start();
                                }
                            }
                        }
                    }
                    //Console.WriteLine("Image transferred and queued in memory.");
                }
                catch (Exception ex)
                {
                    Logger.Log($"Error handling transferred image: {ex.Message}");
                    // Signal completion on fatal error so main thread won't block forever
                    scanCompleted.Set();
                }
            };

            session.TransferError += (s, e) =>
            {
                Logger.Log($"Transfer error: {e.Exception?.Message ?? "Unknown"}");
                // Signal completion on error
                scanCompleted.Set();
            };

            session.TransferCanceled += (s, e) =>
            {
                msg = "Transfer canceled by user.";
                Logger.Log(msg);
                scanCompleted.Set();
            };

            session.SourceDisabled += (s, e) =>
            {
                msg = "Source disabled.";
                //  Log(msg);
                scanCompleted.Set();
            };

            var rc = session.Open();
            if (rc != ReturnCode.Success)
            {
                msg = "Failed to open TWAIN session.";
                Logger.Log(msg);
                transferTimer?.Dispose();
                return msg;
            }

            DataSource ds = null;
            try
            {
                var sources = session.ToList();
                if (sources.Count == 0)
                {
                    msg = "No TWAIN sources found.";
                    //Log(msg);
                    return msg;
                }

                ds = sources[a_sourceIndex];
                ds.Open();

                // Capture scanner identification info
                ScannerName = ds.Name ?? string.Empty;
                ScannerManufacturer = ds.Manufacturer ?? string.Empty;
                ScannerModel = ds.ProductFamily ?? string.Empty;
                Logger.Log($"Scanner: {ScannerName}, Manufacturer: {ScannerManufacturer}, Model: {ScannerModel}");

                if (!ds.IsOpen)
                {
                    msg = "DataSource failed to open.";
                    //Log(msg);
                    return msg;
                }

                SetCapabilties(ds, a_Feeder: a_Feeder, a_Duplex: a_Duplex, a_color: a_color,
                    a_resolution: a_resolution,
                    a_PageWidth: a_PageWidth, a_PageHeight: a_PageHeight);

                // Prepare for acquisition
                scanCompleted.Reset();
                lock (transferLock)
                {
                    transferTimer.Stop();
                    pages.Clear();
                }

                // Start acquisition (provide a message window handle if your driver requires one)
                ds.Enable(SourceEnableMode.NoUI, false, IntPtr.Zero);

                Logger.Log("Waiting for scan to complete...");
                // Wait until the debounce timer elapses after last DataTransferred or TransferError sets event
                scanCompleted.WaitOne();

                // Save all pages to disk (main thread)
                lock (transferLock)
                {
                    if (pages.Count == 0)
                    {
                        msg = msg + " No pages were scanned. pages.Count = 0.";
                        //Log(msg.Trim());
                        return msg;
                    }

                    //// todo debug rng
                    //for (int i = 0; i < pages.Count; i++)
                    //{
                    //    var bmpFile = Path.ChangeExtension(a_fullPathPdf, $"-{i}.bmp");
                    //    pages[i].Save(bmpFile);
                    //}

                    // Create PDF from scanned pages
                    PdfCreator.CreatePdfFromBitmaps(pages, a_fullPathPdf);
                }

                // Close source after transfer completes
                ds.Close();
            }
            finally
            {
                // Dispose all Bitmaps in pages list (important for cancel/error paths)
                lock (transferLock)
                {
                    foreach (var bmp in pages)
                    {
                        try { bmp?.Dispose(); }
                        catch { /* ignore dispose errors */ }
                    }
                    pages.Clear();
                }

                // Close data source if open
                try { if (ds != null && ds.IsOpen) ds.Close(); }
                catch { /* ignore */ }

                session.Close();
                transferTimer?.Dispose();
            }
            //Log("Successful scan completed.");
            msg = "";   // Empty msg indicates success.
            return msg;
        }


        /// <summary>
        /// Comma seaprated index and description of scanners. 
        /// </summary>
        /// <returns>Returns in one string the list of available scanners.</returns>
        public string GetAvailableScanners()
        {
            short scanIndex = -1;
            string scannerListString = string.Empty;
            List<string> scannerList = new List<string>();
            var appId = TWIdentity.CreateFromAssembly(DataGroups.Image, Assembly.GetExecutingAssembly());
            var session = new TwainSession(appId);
            try
            {
                session.Open();
                var sources = session.ToList();
                scannerList = sources.Select(s => s.Name).ToList();
                foreach (var scanner in scannerList)
                {
                    scanIndex++;
                    scannerListString += $"{scanIndex}={scanner}\r\n";
                }
            }
            catch (Exception)
            {
                //throw;
                // Swallow exceptions
            }
            finally
            {
                session.Close();
            }
            return scannerListString;
        }

        public string Greetings()
        {
            return "Greetings, Programs!\ncolor:bw,gray,color\nresolution:low,medium,high";
        }
    }
}
