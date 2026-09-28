// Title: Combine multiple Excel files into one workbook while preserving each sheet’s page‑setup settings using Aspose.Cells for .NET
// AI Prompts: Write a C# routine that accepts an array of .xlsx file paths, merges them into a single Workbook, and copies the PageSetup of every source Worksheet to the corresponding sheet in the merged workbook using Aspose.Cells. | Implement a helper method that transfers all PageSetup properties (orientation, paper size, margins, print area, etc.) from a source Worksheet.PageSetup object to a target Worksheet.PageSetup object in Aspose.Cells. | Add robust error handling that skips missing source files, logs processing errors, and creates the destination folder automatically before saving the merged workbook.
// Common Searches: how to merge several .xlsx files and keep original print margins with Aspose.Cells C# | copy worksheet page setup when adding sheets to a merged workbook in .NET | C# Aspose.Cells merge workbooks preserve orientation and paper size | skip missing Excel files during workbook merge using Aspose.Cells | create output directory automatically when saving merged workbook Aspose.Cells
// Tags: Aspose.Cells merge workbooks with page‑setup copy | C# copy worksheet print layout between workbooks | preserve Excel sheet margins during merge | handle missing source files Aspose.Cells | auto‑create output folder before saving workbook

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookMergeApp
{
    // Provides a C# example that merges an array of Excel files into a single workbook using Aspose.Cells, copies each source worksheet’s PageSetup (orientation, paper size, margins, print area, etc.) to the newly added sheet, skips missing files, creates the output directory if needed, and saves the combined workbook.
    public class WorkbookMerger
    {
        /// <param name="sourceFiles">Array of full paths to source Excel files.</param>
        /// <param name="outputFile">Full path for the merged workbook to be saved.</param>
        public static void MergeWorkbooks(string[] sourceFiles, string outputFile)
        {
            if (sourceFiles == null || sourceFiles.Length == 0)
                throw new ArgumentException("No source files provided.");

            if (string.IsNullOrWhiteSpace(outputFile))
                throw new ArgumentException("Output file path is required.");

            // Create an empty workbook that will hold the merged result.
            Workbook mergedWorkbook = new Workbook();

            // Remove the default empty sheet that Aspose.Cells creates.
            if (mergedWorkbook.Worksheets.Count > 0)
                mergedWorkbook.Worksheets.RemoveAt(0);

            // Iterate through each source file.
            foreach (string srcPath in sourceFiles)
            {
                if (!File.Exists(srcPath))
                {
                    Console.WriteLine($"Source file not found: {srcPath}");
                    continue; // skip missing file
                }

                try
                {
                    // Load the source workbook.
                    Workbook srcWorkbook = new Workbook(srcPath);

                    // Loop through all worksheets in the source workbook.
                    foreach (Worksheet srcSheet in srcWorkbook.Worksheets)
                    {
                        // Add a copy of the source worksheet to the merged workbook.
                        int newIndex = mergedWorkbook.Worksheets.AddCopy(srcSheet.Name);

                        // Retrieve the newly added worksheet.
                        Worksheet destSheet = mergedWorkbook.Worksheets[newIndex];

                        // Copy page‑setup settings from source to destination.
                        CopyPageSetup(srcSheet.PageSetup, destSheet.PageSetup);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{srcPath}': {ex.Message}");
                }
            }

            // Ensure output directory exists.
            string outDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                Directory.CreateDirectory(outDir);

            // Save the merged workbook to the specified output file.
            try
            {
                mergedWorkbook.Save(outputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save merged workbook: {ex.Message}");
                throw;
            }
        }

        private static void CopyPageSetup(PageSetup source, PageSetup target)
        {
            target.Orientation = source.Orientation;
            target.PaperSize = source.PaperSize;
            target.FitToPagesTall = source.FitToPagesTall;
            target.FitToPagesWide = source.FitToPagesWide;
            target.TopMargin = source.TopMargin;
            target.BottomMargin = source.BottomMargin;
            target.LeftMargin = source.LeftMargin;
            target.RightMargin = source.RightMargin;
            target.HeaderMargin = source.HeaderMargin;
            target.FooterMargin = source.FooterMargin;
            target.CenterHorizontally = source.CenterHorizontally;
            target.CenterVertically = source.CenterVertically;
            target.PrintArea = source.PrintArea;
            target.PrintTitleColumns = source.PrintTitleColumns;
            target.PrintTitleRows = source.PrintTitleRows;
            target.BlackAndWhite = source.BlackAndWhite;
            target.PrintGridlines = source.PrintGridlines;
            target.PrintHeadings = source.PrintHeadings;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage – adjust paths as needed.
            string[] files = {
                @"C:\Data\Report1.xlsx",
                @"C:\Data\Report2.xlsx",
                @"C:\Data\Report3.xlsx"
            };
            string output = @"C:\Data\MergedReport.xlsx";

            try
            {
                WorkbookMerger.MergeWorkbooks(files, output);
                Console.WriteLine($"Merged workbook saved to: {output}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
