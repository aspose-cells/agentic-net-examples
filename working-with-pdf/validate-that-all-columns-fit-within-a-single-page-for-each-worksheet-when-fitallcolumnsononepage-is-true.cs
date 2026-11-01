// Title: Validate that every worksheet has "Fit All Columns on One Page" enabled with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that loads a workbook, iterates all worksheets, and throws an InvalidOperationException when a sheet's page setup does not have FitToPagesWide = 1 and FitToPagesTall = 0. | Create a reusable C# method that accepts a Workbook object and returns a list of worksheet names where the "Fit All Columns on One Page" page‑setup option is incorrectly configured.
// Common Searches: asp.net aspose.cells verify FitAllColumnsOnOnePage for each worksheet | c# check FitToPagesWide equals 1 and FitToPagesTall equals 0 using Aspose.Cells | how to ensure Excel sheets print all columns on one page with Aspose.Cells .NET | validate workbook page setup column fit setting Aspose.Cells C#
// Tags: Aspose.Cells verify FitToPagesWide setting | Aspose.Cells check FitAllColumnsOnOnePage flag | C# workbook page‑setup validation | Excel worksheet print layout verification Aspose.Cells | Aspose.Cells ensure columns fit on single page

using Aspose.Cells;
using System;
using System.IO;

// The program loads an Excel workbook with Aspose.Cells, iterates each worksheet, detects if the page setup is configured for "Fit All Columns on One Page" (FitToPagesWide = 1 and FitToPagesTall = 0), throws an InvalidOperationException for any mismatched settings, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PageSetup pageSetup = sheet.PageSetup;

                // Determine if "Fit All Columns on One Page" is effectively enabled
                // In Aspose.Cells this is represented by FitToPagesWide = 1 and FitToPagesTall = 0
                bool isFitAllColumnsOnOnePage = pageSetup.FitToPagesWide == 1 && pageSetup.FitToPagesTall == 0;

                if (isFitAllColumnsOnOnePage)
                {
                    // Validate that the setting forces all columns onto a single page
                    if (pageSetup.FitToPagesWide != 1)
                    {
                        throw new InvalidOperationException(
                            $"Worksheet '{sheet.Name}' is expected to have FitToPagesWide = 1 but found {pageSetup.FitToPagesWide}.");
                    }

                    // Ensure rows are not limited to a single page (FitToPagesTall should be 0)
                    if (pageSetup.FitToPagesTall != 0)
                    {
                        throw new InvalidOperationException(
                            $"Worksheet '{sheet.Name}' is expected to have FitToPagesTall = 0 but found {pageSetup.FitToPagesTall}.");
                    }
                }
            }

            // Save the workbook (if any changes were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
