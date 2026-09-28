// Title: How to clone the page setup from a template worksheet to a new worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads a template workbook, reads its first worksheet's PageSetup (paper size, orientation, margins, print area, fit‑to‑pages), creates a new worksheet, applies the same PageSetup values, and saves the result as Result.xlsx using Aspose.Cells. | Generate a program that checks for the existence of Template.xlsx, copies all printable layout settings from the source sheet to a newly added sheet named 'ClonedPageSetup', and handles any missing properties gracefully with Aspose.Cells. | Create a C# console application that duplicates the page‑setup configuration (including margins, paper size, orientation, and fit‑to‑page options) from a template worksheet to another worksheet in a separate workbook, then saves the workbook.
// Common Searches: aspnet copy worksheet page setup properties from one Excel file to another using Aspose.Cells | c# clone print area, margins and orientation of a worksheet with Aspose.Cells | how to transfer page layout settings between worksheets in Aspose.Cells for .NET | Aspose.Cells example to duplicate page setup from template sheet to new sheet
// Tags: Aspose.Cells copy worksheet page setup | C# clone Excel worksheet print settings | Aspose.Cells transfer page layout properties | duplicate margins orientation paper size Aspose.Cells | copy print area between worksheets C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads Template.xlsx, extracts the first worksheet's PageSetup settings (paper size, orientation, margins, print area, and fit‑to‑page options), creates a new workbook with a sheet named 'ClonedPageSetup', applies the extracted settings to the new sheet, and saves the workbook as Result.xlsx while handling missing files and potential copy errors.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the template workbook
            const string templatePath = "Template.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Error: Template file not found at '{templatePath}'.");
                return;
            }

            // Load the template workbook containing the source worksheet
            Workbook templateWorkbook = new Workbook(templatePath);
            Worksheet templateSheet = templateWorkbook.Worksheets[0];

            // Create a new workbook that will hold the newly created sheet
            Workbook newWorkbook = new Workbook();
            Worksheet newSheet = newWorkbook.Worksheets[0];
            newSheet.Name = "ClonedPageSetup";

            // Clone the page setup from the template sheet to the new sheet
            try
            {
                newSheet.PageSetup.PaperSize = templateSheet.PageSetup.PaperSize;
                newSheet.PageSetup.Orientation = templateSheet.PageSetup.Orientation;
                newSheet.PageSetup.PrintArea = templateSheet.PageSetup.PrintArea;
                newSheet.PageSetup.FitToPagesWide = templateSheet.PageSetup.FitToPagesWide;
                newSheet.PageSetup.FitToPagesTall = templateSheet.PageSetup.FitToPagesTall;
                newSheet.PageSetup.BottomMargin = templateSheet.PageSetup.BottomMargin;
                newSheet.PageSetup.TopMargin = templateSheet.PageSetup.TopMargin;
                newSheet.PageSetup.LeftMargin = templateSheet.PageSetup.LeftMargin;
                newSheet.PageSetup.RightMargin = templateSheet.PageSetup.RightMargin;

                // Header/footer properties are not available in this version of Aspose.Cells.
                // If needed, they can be set using the appropriate API for the used version.
            }
            catch (Exception copyEx)
            {
                Console.WriteLine($"Warning: Failed to copy some page setup properties. {copyEx.Message}");
            }

            // Save the resulting workbook
            const string resultPath = "Result.xlsx";
            newWorkbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
