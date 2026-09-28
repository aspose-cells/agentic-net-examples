// Title: Remove printer settings and print areas from all worksheets in an Excel file using Aspose.Cells for .NET so the workbook opens without a default‑printer prompt
// AI Prompts: Generate C# code with Aspose.Cells that loads an .xlsx file, iterates through each worksheet, sets PageSetup.PrinterSettings to null, clears PrintArea, PrintTitleRows, and PrintTitleColumns, then saves the modified workbook. | Show how to confirm that a workbook saved after clearing its printer configuration opens in Excel without displaying the "Select a printer" dialog. | Provide a robust C# snippet that checks for the input file, handles exceptions, and removes all printer‑related PageSetup options from a workbook using Aspose.Cells.
// Common Searches: Aspose.Cells how to delete printer settings from an existing workbook in C# | C# remove print area and printer configuration from Excel file to avoid default printer dialog | prevent Excel from asking for a printer when opening a file generated with Aspose.Cells | clear PageSetup.PrinterSettings for all worksheets using Aspose.Cells .NET | sample code to reset worksheet print titles and printer settings in Aspose.Cells
// Tags: Aspose.Cells clear worksheet printer settings | PageSetup.PrinterSettings null C# | reset worksheet print area Aspose.Cells | clear print title rows and columns Aspose.Cells | prevent Excel default printer dialog Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads input.xlsx, iterates through every worksheet, nullifies PageSetup.PrinterSettings, empties PrintArea, PrintTitleRows, and PrintTitleColumns, then saves the result as output.xlsx while handling missing files and runtime exceptions.
class RemovePrinterSettings
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and clear printer settings
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PageSetup pageSetup = sheet.PageSetup;

                // Clear stored printer settings (byte[]) if present
                if (pageSetup.PrinterSettings != null && pageSetup.PrinterSettings.Length > 0)
                {
                    pageSetup.PrinterSettings = null;
                }

                // Remove other printer‑related options that may cause prompts
                pageSetup.PrintArea = string.Empty;          // Remove any defined print area
                pageSetup.PrintTitleRows = string.Empty;     // Remove title rows
                pageSetup.PrintTitleColumns = string.Empty;  // Remove title columns
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
