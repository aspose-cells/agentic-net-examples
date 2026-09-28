// Title: Clone an Excel worksheet in C# with Aspose.Cells, preserving page‑setup settings and background image
// AI Prompts: Write C# code that uses Aspose.Cells to duplicate a worksheet, copy its BackgroundImage byte array, and transfer all PageSetup properties to the new sheet. | Create a reusable method in C# that clones a worksheet via AddCopy, copies orientation, margins, print area, and background image without sharing references. | Generate a complete example that loads a workbook, clones a specific sheet, preserves page‑setup configuration and background image, and saves the result using Aspose.Cells.
// Common Searches: aspnet clone worksheet preserve page setup Aspose.Cells | copy background image of Excel sheet using Aspose.Cells C# | how to keep margins and print area when duplicating a worksheet in .NET | Aspose.Cells AddCopy retain background image stream | duplicate Excel sheet with same page setup settings C#
// Tags: Aspose.Cells sheet copy | background image data copy | page‑setup property transfer | AddCopy sheet cloning method | preserve sheet margins and print area

using System;
using System.IO;
using Aspose.Cells;

// The example loads a source workbook, clones the specified worksheet using the AddCopy method, copies every relevant PageSetup property (orientation, margins, print area, etc.) to the cloned sheet, duplicates the background image byte array to avoid reference sharing, and saves the workbook with the new cloned worksheet.
class WorksheetCloneExample
{
    static void Main()
    {
        try
        {
            string sourcePath = "SourceWorkbook.xlsx";

            // Ensure the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Get the source worksheet (by name)
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"];
            if (sourceSheet == null)
            {
                Console.WriteLine("Source worksheet 'Sheet1' not found.");
                return;
            }

            int sourceIndex = workbook.Worksheets.IndexOf(sourceSheet);

            // Clone the worksheet
            int clonedIndex = workbook.Worksheets.AddCopy(sourceIndex);
            Worksheet clonedSheet = workbook.Worksheets[clonedIndex];

            // ----- Preserve original page‑setup configuration -----
            PageSetup srcSetup = sourceSheet.PageSetup;
            PageSetup dstSetup = clonedSheet.PageSetup;

            dstSetup.Orientation = srcSetup.Orientation;
            dstSetup.PaperSize = srcSetup.PaperSize;
            dstSetup.TopMargin = srcSetup.TopMargin;
            dstSetup.BottomMargin = srcSetup.BottomMargin;
            dstSetup.LeftMargin = srcSetup.LeftMargin;
            dstSetup.RightMargin = srcSetup.RightMargin;
            dstSetup.HeaderMargin = srcSetup.HeaderMargin;
            dstSetup.FooterMargin = srcSetup.FooterMargin;
            dstSetup.CenterHorizontally = srcSetup.CenterHorizontally;
            dstSetup.CenterVertically = srcSetup.CenterVertically;
            dstSetup.PrintArea = srcSetup.PrintArea;
            dstSetup.PrintTitleRows = srcSetup.PrintTitleRows;
            dstSetup.PrintTitleColumns = srcSetup.PrintTitleColumns;
            // Add any additional PageSetup properties you need to preserve

            // ----- Copy background image -----
            if (sourceSheet.BackgroundImage != null && sourceSheet.BackgroundImage.Length > 0)
            {
                // Clone the byte array to avoid sharing the same reference
                byte[] bgImage = new byte[sourceSheet.BackgroundImage.Length];
                Array.Copy(sourceSheet.BackgroundImage, bgImage, bgImage.Length);
                clonedSheet.BackgroundImage = bgImage;
            }

            // Save the workbook with the cloned worksheet
            string outputPath = "ClonedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
