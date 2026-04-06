using NTwain;
using NTwain.Data;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
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

            // Check TWAIN platform support; fall back to WIA if not available
            if (!PlatformInfo.Current.IsSupported)
            {
                //var wia = new WiaScanner();
                //return wia.ScanToPdf(a_sourceIndex, a_fullPathPdf,
                //    a_Feeder, a_Duplex,
                //    a_color, a_resolution,
                //    a_PageWidth, a_PageHeight);
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

            //Log($"Parameters: a_sourceIndex={a_sourceIndex},\na_fullPathPdf={a_fullPathPdf},\na_Feeder={a_Feeder},\na_Duplex={a_Duplex},\na_color={a_color},\na_resolution={a_resolution},\na_PageWidth={a_PageWidth},\na_PageHeight={a_PageHeight}");
            //if (a_Feeder) { Log("Feeder is true."); }  else Log("Feeder is FALSE,");
            //if (a_Duplex) { Log("Duplex is true."); } else Log("Dulex is FALSE.");

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
                    //Log($"Error handling transferred image: {ex.Message}");
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

                //Log("Waiting for scan to complete...");
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




                    // Create a new PDF document
                    #region ⚠️ Important Note ⚠️
                    // Using older version of PdfSharp (1.50.5147) from NuGet because we only need minimum features.
                    // Ignore the warnings about this version being obsolete.
                    #endregion
                    using (PdfDocument pdf = new PdfDocument())
                    {
                        pdf.Info.Title = "Scanned Document";
                        // Add each scanned page as an image to the PDF
                        for (int i = 0; i < pages.Count; i++)
                        {
                            var page = pages[i];
                            if (page == null || page.Size == Size.Empty) continue;

                            // Add a new page to the document
                            var pdfPage = pdf.AddPage();
                            pdfPage.Size = PdfSharp.PageSize.Letter;
                            pdfPage.Orientation = PdfSharp.PageOrientation.Portrait;

                            using (XGraphics gfx = XGraphics.FromPdfPage(pdfPage))
                            {
                                // Draw the image on the PDF page
                                using (var ms = new MemoryStream())
                                {
                                    // Save the bitmap to the stream as a JPEG file
                                    page.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                                    ms.Position = 0;

                                    // Load the image from the stream
                                    using (XImage img = XImage.FromStream(ms))
                                    {
                                        // Calculate the scaling to fit the image into the PDF page.
                                        double xScale = pdfPage.Width / img.Width;
                                        double yScale = pdfPage.Height / img.Height;

                                        double scale = Math.Min(xScale, yScale);
                                        //Log($"scale: {scale}");

                                        // Calculate the position to center the image on the PDF page
                                        double x = (pdfPage.Width - img.Width * scale) / 2;
                                        double y = (pdfPage.Height - img.Height * scale) / 2;

                                        // Draw the image with the calculated size and position
                                        gfx.DrawImage(img, x, y, img.PixelWidth * scale, img.PixelHeight * scale);
                                    }
                                }
                            }

                            // Mark as disposed by setting to null (actual dispose in finally)
                            pages[i] = null;
                            page.Dispose();
                        }

                        // Save the document to a file
                        pdf.Save(a_fullPathPdf);
                        Console.WriteLine($"Saved document to {a_fullPathPdf}");
                    }
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

        //private void Log(string message)
        //{
        //    Helpers.Log(logFile, message);
        //}


    }
}
