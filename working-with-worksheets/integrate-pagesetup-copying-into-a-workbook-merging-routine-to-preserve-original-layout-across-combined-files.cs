// Title: Merge multiple Excel workbooks in C# while copying worksheet PageSetup settings with Aspose.Cells
// Common Searches: c# merge several .xlsx files and keep original page setup using Aspose.Cells | how to copy worksheet print settings when combining Excel workbooks in .NET | Aspose.Cells preserve margins and orientation during workbook merge | copy page setup from source to destination worksheet while merging workbooks in C# | merge Excel workbooks without losing print area configuration Aspose.Cells
// Tags: workbook merge page setup cloning Aspose.Cells | copy worksheet print options C# | retain Excel print margins during merge | transfer page orientation Aspose.Cells | duplicate page setup across merged worksheets

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorkbookMergeExample
{
    // The example loads a list of source .xlsx files, adds copies of each worksheet to a new destination workbook, and uses a helper method to copy all PageSetup properties (paper size, orientation, zoom, margins, print area, gridlines, headings, comments, centering, etc.) from the source worksheet to the newly added worksheet before saving the merged workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // List of source workbook file paths to be merged
            List<string> sourceFiles = new List<string>
            {
                @"C:\Data\Source1.xlsx",
                @"C:\Data\Source2.xlsx",
                @"C:\Data\Source3.xlsx"
            };

            // Destination workbook where all sheets will be merged
            Workbook destination = new Workbook(); // create-workbook rule
            // Remove the default empty sheet created by Aspose.Cells
            destination.Worksheets.Clear();

            foreach (string filePath in sourceFiles)
            {
                try
                {
                    // Ensure the source file exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Source file not found: {filePath}");
                        continue;
                    }

                    // Load each source workbook
                    Workbook source = new Workbook(filePath); // load-workbook rule

                    // Iterate through all worksheets in the source workbook
                    foreach (Worksheet srcSheet in source.Worksheets)
                    {
                        // Add a copy of the source worksheet to the destination workbook
                        int newIndex = destination.Worksheets.AddCopy(srcSheet.Name); // add-copy rule

                        // Retrieve the newly added worksheet
                        Worksheet destSheet = destination.Worksheets[newIndex];

                        // Copy page‑setup settings from source to destination worksheet
                        CopyPageSetup(srcSheet.PageSetup, destSheet.PageSetup);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            try
            {
                // Ensure the output directory exists
                string outputPath = @"C:\Data\MergedResult.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the merged workbook
                destination.Save(outputPath); // save-workbook rule
                Console.WriteLine($"Merged workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving merged workbook: {ex.Message}");
            }
        }

        /// <param name="source">The source PageSetup.</param>
        /// <param name="dest">The destination PageSetup.</param>
        private static void CopyPageSetup(PageSetup source, PageSetup dest)
        {
            // Basic layout settings
            dest.PaperSize = source.PaperSize;
            dest.Orientation = source.Orientation;
            dest.Zoom = source.Zoom;
            dest.FitToPagesTall = source.FitToPagesTall;
            dest.FitToPagesWide = source.FitToPagesWide;

            // Margins (in points)
            dest.TopMargin = source.TopMargin;
            dest.BottomMargin = source.BottomMargin;
            dest.LeftMargin = source.LeftMargin;
            dest.RightMargin = source.RightMargin;
            dest.HeaderMargin = source.HeaderMargin;
            dest.FooterMargin = source.FooterMargin;

            // Print options
            dest.PrintGridlines = source.PrintGridlines;
            dest.PrintHeadings = source.PrintHeadings;
            dest.PrintComments = source.PrintComments;
            dest.PrintArea = source.PrintArea;

            // Centering options
            dest.CenterHorizontally = source.CenterHorizontally;
            dest.CenterVertically = source.CenterVertically;

            // Additional settings can be added here as needed
        }
    }
}
