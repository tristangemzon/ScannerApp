using NTwain;
using NTwain.Data;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Controls;
using System.Windows.Interop;

namespace ScannerApp
{
    public partial class TwainScanner
    {
        internal string SetCapabilties(DataSource ds,
                    bool a_Feeder, bool a_Duplex,
                    string a_color,
                    string a_resolution,
                    int a_PageWidth, int a_PageHeight)
        {
            string msg = string.Empty;

            try
            {
                // Do not scan blank pages.
                if (ds.Capabilities.ICapAutoDiscardBlankPages != null && ds.Capabilities.ICapAutoDiscardBlankPages.CanSet)
                    ds.Capabilities.ICapAutoDiscardBlankPages.SetValue(BlankPage.Invalid);
            }
            catch (Exception ex)
            {
                Logger.Log($"Warning: couldn't set blank-page discard capability: {ex.Message}");
            }

            if (a_Feeder == true && ds.Capabilities.CapFeederEnabled.CanSet)
                ds.Capabilities.CapFeederEnabled.SetValue(BoolType.True);

            if (a_Duplex == true && ds.Capabilities.CapDuplexEnabled.CanSet)
                ds.Capabilities.CapDuplexEnabled.SetValue(BoolType.True);

            if (ds.Capabilities.ICapSupportedSizes != null && ds.Capabilities.ICapSupportedSizes.CanSet)
            {
                var paperSize = TwainPaperSizeDetector.Detect((int)a_PageWidth, (int)a_PageHeight);
                if (paperSize != null)
                    ds.Capabilities.ICapSupportedSizes.SetValue((SupportedSize)paperSize);
                var gv = ds.Capabilities.ICapSupportedSizes.GetCurrent();
            }

            // Resolution
            float dpi;
            switch (a_resolution)
            {
                case "low":
                    dpi = 100f;
                    break;
                case "medium":
                    dpi = 200f;
                    break;
                case "high":
                    dpi = 300f;
                    break;
                default:
                    dpi = 200f;
                    break;
            }
            if (ds.Capabilities.ICapXResolution.CanSet)
                ds.Capabilities.ICapXResolution.SetValue(dpi);
            if (ds.Capabilities.ICapYResolution.CanSet)
                ds.Capabilities.ICapYResolution.SetValue(dpi);

            // Step 1: Set PixelType
            PixelType pt;
            switch (a_color)
            {
                case "bw":
                    pt = PixelType.BlackWhite;
                    break;
                case "gray":
                    pt = PixelType.Gray;
                    break;
                case "color":
                    pt = PixelType.RGB;
                    break;
                default:
                    pt = PixelType.RGB;
                    break;
            }
            ds.Capabilities.ICapPixelType.SetValue(pt);

            // Step 2: Query supported bit depths
            var supported = ds.Capabilities.ICapBitDepth.GetValues();
            int chosen = 0;

            // Step 3: Choose best bit depth based on pixel type
            if (pt == PixelType.BlackWhite)
            {
                // Prefer 1-bit
                chosen = supported.Contains(1) ? 1 : supported.FirstOrDefault();
            }
            else if (pt == PixelType.Gray)
            {
                // Prefer 8-bit grayscale
                if (supported.Contains(8))
                    chosen = 8;
                else if (supported.Contains(16))
                    chosen = 16;
                else
                    chosen = supported.FirstOrDefault();
            }
            else if (pt == PixelType.RGB)
            {
                // Prefer 8 bits per channel (24 total)
                if (supported.Contains(24))
                    chosen = 24;
                else if (supported.Contains(8))
                    chosen = 8; // Some scanners report 8 instead of 24
                else
                    chosen = supported.FirstOrDefault();
            }

            // Step 4: Apply bit depth
            if (chosen != 0)
                ds.Capabilities.ICapBitDepth.SetValue(chosen);


            // working settings example - DO NOT DELETE
            //if (ds.Capabilities.CapFeederEnabled.CanSet)
            //    ds.Capabilities.CapFeederEnabled.SetValue(BoolType.True);
            //if (ds.Capabilities.CapDuplexEnabled.CanSet)
            //    ds.Capabilities.CapDuplexEnabled.SetValue(BoolType.True);
            //if (ds.Capabilities.ICapSupportedSizes != null && ds.Capabilities.ICapSupportedSizes.CanSet)
            //    ds.Capabilities.ICapSupportedSizes.SetValue(SupportedSize.USLetter);
            //if (ds.Capabilities.ICapPixelType.CanSet)
            //    ds.Capabilities.ICapPixelType.SetValue(PixelType.Gray);
            //if (ds.Capabilities.ICapBitDepth.CanSet)
            //    ds.Capabilities.ICapBitDepth.SetValue(8);           // 1, 8, 24. bw, gray and color respectively.
            //if (ds.Capabilities.ICapXResolution.CanSet)
            //    ds.Capabilities.ICapXResolution.SetValue(300);
            //if (ds.Capabilities.ICapYResolution.CanSet)
            //    ds.Capabilities.ICapYResolution.SetValue(300);
            return msg;
        }
    }


    //public enum PaperSize
    //{
    //    Letter,
    //    Legal,
    //    A4,
    //    A5,
    //    A3,
    //    B5,
    //    Custom
    //}

    internal static class TwainPaperSizeDetector
    {
        // TWAIN uses 1/1000 inch units
        private static double ToInches(int twainUnits) => twainUnits / 1000.0;

        private static bool IsClose(double a, double b, double tolerance = 0.05)
            => Math.Abs(a - b) <= tolerance;

        public static SupportedSize? Detect(int widthUnits, int heightUnits)
        {
            double w = ToInches(widthUnits);
            double h = ToInches(heightUnits);

            // Normalize orientation (Letter can be 8.5x11 or 11x8.5)
            double min = Math.Min(w, h);
            double max = Math.Max(w, h);

            // Letter: 8.5 x 11
            if (IsClose(min, 8.5) && IsClose(max, 11.0))
                return SupportedSize.USLetter;

            // Legal: 8.5 x 14
            if (IsClose(min, 8.5) && IsClose(max, 14.0))
                return SupportedSize.USLegal;

            // A4: 8.27 x 11.69
            if (IsClose(min, 8.27) && IsClose(max, 11.69))
                return SupportedSize.A4;

            // A5: 5.83 x 8.27
            if (IsClose(min, 5.83) && IsClose(max, 8.27))
                return SupportedSize.A5;

            // A3: 11.69 x 16.54
            if (IsClose(min, 11.69) && IsClose(max, 16.54))
                return SupportedSize.A3;

            // B5: 7.17 x 10.12
            if (IsClose(min, 7.17) && IsClose(max, 10.12))
                return SupportedSize.IsoB5;

            return null;
        }
    }

}
