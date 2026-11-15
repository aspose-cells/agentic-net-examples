// Title: Copy full page‑setup settings from one Excel worksheet to another using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells copy worksheet page setup to another sheet | How to duplicate print area and margins between worksheets with Aspose.Cells | Transfer Excel page layout (orientation, paper size) using Aspose.Cells .NET | Clone worksheet print options programmatically in C# with Aspose.Cells | Copy page setup properties from source to destination worksheet Aspose.Cells example
// Tags: Aspose.Cells copy worksheet page setup C# | duplicate Excel print area margins Aspose.Cells | clone worksheet print options .NET | transfer page layout settings between sheets Aspose.Cells | copy page setup properties Aspose.Cells API | worksheet page setup cloning C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, selects the first worksheet as the source, creates or retrieves a second worksheet as the destination, and copies every page‑setup property—including orientation, paper size, fit‑to‑pages, zoom, print area, title rows/columns, margins, centering, and print options—from the source to the destination worksheet before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get source worksheet (first sheet)
                Worksheet sourceSheet = workbook.Worksheets[0];
                Worksheet destSheet;

                // Create or get destination sheet (second sheet)
                if (workbook.Worksheets.Count > 1)
                {
                    destSheet = workbook.Worksheets[1];
                }
                else
                {
                    // Worksheets.Add() may return the index of the new sheet in some versions
                    int newIndex = workbook.Worksheets.Add();
                    destSheet = workbook.Worksheets[newIndex];
                }

                // Copy page‑setup configuration
                CopyPageSetup(sourceSheet, destSheet);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Copies the entire page‑setup configuration from one worksheet to another
        static void CopyPageSetup(Worksheet sourceWorksheet, Worksheet destinationWorksheet)
        {
            try
            {
                // Get the PageSetup objects of both worksheets
                PageSetup srcSetup = sourceWorksheet.PageSetup;
                PageSetup destSetup = destinationWorksheet.PageSetup;

                // General page layout
                destSetup.Orientation = srcSetup.Orientation;
                destSetup.PaperSize = srcSetup.PaperSize;
                destSetup.FitToPagesWide = srcSetup.FitToPagesWide;
                destSetup.FitToPagesTall = srcSetup.FitToPagesTall;
                destSetup.Zoom = srcSetup.Zoom;

                // Print area and titles
                destSetup.PrintArea = srcSetup.PrintArea;
                destSetup.PrintTitleRows = srcSetup.PrintTitleRows;
                destSetup.PrintTitleColumns = srcSetup.PrintTitleColumns;

                // Margins (in points)
                destSetup.LeftMargin = srcSetup.LeftMargin;
                destSetup.RightMargin = srcSetup.RightMargin;
                destSetup.TopMargin = srcSetup.TopMargin;
                destSetup.BottomMargin = srcSetup.BottomMargin;
                destSetup.HeaderMargin = srcSetup.HeaderMargin;
                destSetup.FooterMargin = srcSetup.FooterMargin;

                // Centering options
                destSetup.CenterHorizontally = srcSetup.CenterHorizontally;
                destSetup.CenterVertically = srcSetup.CenterVertically;

                // Print options
                destSetup.BlackAndWhite = srcSetup.BlackAndWhite;
                destSetup.PrintDraft = srcSetup.PrintDraft;
                destSetup.PrintGridlines = srcSetup.PrintGridlines;
                destSetup.PrintHeadings = srcSetup.PrintHeadings;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error copying page setup: {ex.Message}");
            }
        }
    }
}
