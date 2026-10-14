// Title: Remove printer‑related page‑setup settings from all worksheets in an XLSX file using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that opens an existing XLSX, clears the print area, title rows/columns, fit‑to‑page and zoom settings on every worksheet, and saves the cleaned workbook as a new file. | Create a reusable C# method that accepts an input Excel path, removes all page‑setup properties (print area, titles, fit‑to‑page, zoom) from each sheet using Aspose.Cells, and returns the modified Workbook object. | Update the sample program to also reset page margins to their default values after clearing printer settings, then save the workbook.
// Common Searches: how to clear printer settings from all worksheets with Aspose.Cells C# | remove print area and title rows from Excel file using Aspose.Cells .NET | reset fit-to-page and zoom for every sheet in an XLSX using Aspose.Cells | save Excel workbook without page‑setup information Aspose.Cells
// Tags: Aspose.Cells remove worksheet printer settings | clear page setup Aspose.Cells C# | reset print area Aspose.Cells workbook | clean Excel file printer configuration Aspose.Cells | save cleaned XLSX Aspose.Cells .NET | iterate worksheets Aspose.Cells page setup

using System;
using System.IO;
using Aspose.Cells;

// The program loads an XLSX workbook, iterates through each worksheet, clears printer‑related page‑setup properties (print area, title rows/columns, fit‑to‑page, zoom), and saves the cleaned workbook as a new XLSX file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "clean.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Clear printer‑related settings from each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Remove print area and title rows/columns
                sheet.PageSetup.PrintArea = null;
                sheet.PageSetup.PrintTitleRows = null;
                sheet.PageSetup.PrintTitleColumns = null;

                // Reset fit‑to‑page settings to defaults
                sheet.PageSetup.FitToPagesTall = 0;
                sheet.PageSetup.FitToPagesWide = 0;

                // Reset zoom to 100%
                sheet.PageSetup.Zoom = 100;
            }

            // Save the cleaned workbook as a new XLSX file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
