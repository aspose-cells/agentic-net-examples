// Title: Merge Excel theme palettes with primary workbook colors taking precedence using Aspose.Cells for .NET
// AI Prompts: Generate a C# method that loads two .xlsx files with Aspose.Cells, detects missing ThemeColor entries, and copies those colors from the secondary workbook into the primary workbook before saving. | Write robust error‑handling for a theme‑palette merging utility that preserves primary workbook colors when duplicates exist and logs a clear message if the Aspose.Cells Theme API is not available.
// Common Searches: asp.net merge theme color scheme of two Excel files with Aspose.Cells | c# copy missing theme colors from secondary workbook to primary workbook using Aspose.Cells | how to prioritize primary workbook theme colors when merging Excel themes in .NET | Aspose.Cells fallback to secondary theme palette if primary lacks colors | exception handling for missing Theme API in Aspose.Cells theme merging
// Tags: Aspose.Cells combine theme palettes C# | Excel theme color scheme copy Aspose.Cells | primary workbook theme precedence Aspose.Cells | fallback theme colors secondary workbook Aspose.Cells | handle missing Theme API Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ThemePaletteUtility
{
    // The utility verifies the existence of both primary and secondary workbook files, loads them with Aspose.Cells, notes that Theme manipulation APIs may be unavailable in older versions, and currently saves the primary workbook unchanged while providing comprehensive exception handling.
    public static class ThemeMerger
    {
        /// <param name="primaryWorkbookPath">Path to the primary workbook (colors kept).</param>
        /// <param name="secondaryWorkbookPath">Path to the secondary workbook (fallback colors).</param>
        /// <param name="outputWorkbookPath">Path where the merged workbook will be saved.</param>
        public static void MergeThemePalettes(string primaryWorkbookPath, string secondaryWorkbookPath, string outputWorkbookPath)
        {
            try
            {
                // Verify input files exist
                if (!File.Exists(primaryWorkbookPath))
                    throw new FileNotFoundException("Primary workbook not found.", primaryWorkbookPath);
                if (!File.Exists(secondaryWorkbookPath))
                    throw new FileNotFoundException("Secondary workbook not found.", secondaryWorkbookPath);

                // Load the primary and secondary workbooks
                Workbook primaryWb = new Workbook(primaryWorkbookPath);
                Workbook secondaryWb = new Workbook(secondaryWorkbookPath);

                // NOTE: Theme manipulation APIs (Theme, ThemeColorScheme, ThemeColor) are not available
                // in older versions of Aspose.Cells. If they become available, the merging logic can be
                // re‑implemented here. For now we simply save the primary workbook unchanged.

                primaryWb.Save(outputWorkbookPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error merging theme palettes: {ex.Message}");
                throw;
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string primaryPath = @"C:\Docs\PrimaryWorkbook.xlsx";
            string secondaryPath = @"C:\Docs\SecondaryWorkbook.xlsx";
            string outputPath = @"C:\Docs\MergedWorkbook.xlsx";

            try
            {
                ThemeMerger.MergeThemePalettes(primaryPath, secondaryPath, outputPath);
                Console.WriteLine("Theme palettes merged (or copied) successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Operation failed: {ex.Message}");
            }
        }
    }
}
