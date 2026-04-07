using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace ScannerApp
{
    /// <summary>
    /// Utility class for creating PDF documents from scanned images.
    /// </summary>
    /// <remarks>
    /// Using older version of PdfSharp (1.50.5147) from NuGet because we only need minimum features.
    /// Ignore the warnings about this version being obsolete.
    /// </remarks>
    public static class PdfCreator
    {
        /// <summary>
        /// Creates a PDF document from a list of bitmap images.
        /// </summary>
        /// <param name="pages">List of bitmap images to include in the PDF. Items are disposed after being added.</param>
        /// <param name="outputPath">Full path for the output PDF file.</param>
        /// <param name="title">Document title (default: "Scanned Document").</param>
        public static void CreatePdfFromBitmaps(List<Bitmap> pages, string outputPath, string title = "Scanned Document")
        {
            using (PdfDocument pdf = new PdfDocument())
            {
                pdf.Info.Title = title;

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
                                double xScale = pdfPage.Width.Point / img.PixelWidth;
                                double yScale = pdfPage.Height.Point / img.PixelHeight;

                                double scale = Math.Min(xScale, yScale);

                                // Calculate the position to center the image on the PDF page
                                double x = (pdfPage.Width.Point - img.PixelWidth * scale) / 2;
                                double y = (pdfPage.Height.Point - img.PixelHeight * scale) / 2;

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
                pdf.Save(outputPath);
                Console.WriteLine($"Saved document to {outputPath}");
            }
        }
    }
}
