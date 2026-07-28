// Title: C# – Save a Workbook with a Specific Printer Paper Size Using Aspose.Cells LightCells API
// Description: Creates a new Workbook, adds sample data, sets the default printer paper size (e.g., A5) via workbook.Settings.PaperSize, and saves the file as XLSX using the LightCells API to preserve the print layout.
// Keywords: Aspose.Cells | LightCells | C# | PaperSize | PaperA5 | Set printer paper size | Workbook.Settings.PaperSize | Save workbook | Excel export | Print layout
// Common Searches: how to set printer paper size in Aspose.Cells C# | save Excel file with A5 paper size using LightCells | Aspose.Cells default paper size before saving | C# example for workbook.Settings.PaperSize | LightCells API print layout configuration
// Developer Intent: Configure the workbook’s default printer paper size and persist it when saving the file.
// Use Cases: Generate sales reports that must print on A5 stationery. | Create batch invoices pre‑configured for A5 printing to streamline mailing. | Produce printable flyers with a fixed paper size before distribution.
// AI Prompts: Show C# code that sets Workbook.Settings.PaperSize to PaperA5 using LightCells before saving. | Provide a LightCells example that saves an Excel workbook with a specific printer paper size and verifies the setting. | Explain how to read the PaperSize property after loading a workbook saved with LightCells.

using System;
using System.IO;
using Aspose.Cells;

// Creates a new Workbook, adds sample data, sets the default printer paper size (e.g., A5) via workbook.Settings.PaperSize, and saves the file as XLSX using the LightCells API to preserve the print layout.
class LightCellsPaperSizeDemo
{
    static void Main()
    {
        try
        {
            // ------------------------------------------------------------
            // 1. Create a workbook and populate it with sample data
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate first row with sample data
            sheet.Cells[0, 0].PutValue("Aspose.Cells LightCells Demo");
            sheet.Cells[0, 1].PutValue(DateTime.Now);

            // ------------------------------------------------------------
            // 2. Set the default printer paper size for the workbook
            //    This influences the print layout when the workbook is rendered
            // ------------------------------------------------------------
            workbook.Settings.PaperSize = PaperSizeType.PaperA5; // Example: A5 size

            // Verify the setting (optional)
            Console.WriteLine("Workbook default paper size: " + workbook.Settings.PaperSize);

            // ------------------------------------------------------------
            // 3. Save the workbook to disk
            // ------------------------------------------------------------
            string outputPath = "LightCells_A5_Output.xlsx";

            // Ensure the directory exists (in case a relative path with folders is used)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved with A5 paper size to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
