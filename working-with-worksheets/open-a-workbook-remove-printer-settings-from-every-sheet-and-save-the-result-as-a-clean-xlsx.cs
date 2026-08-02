// Title: C# – Remove All Worksheet Printer Settings with Aspose.Cells and Save a Clean XLSX
// Description: Loads an existing Excel file (or creates a new workbook if missing), iterates through every worksheet, clears its printer configuration by setting PageSetup.PrinterSettings to null, and saves the result as a clean XLSX file.
// Keywords: Aspose.Cells printer settings | clear worksheet printer configuration | C# remove printer settings Excel | Aspose.Cells PageSetup PrinterSettings null | sanitize Excel workbook .NET | save clean XLSX Aspose | remove page setup printer info | Excel file without printer settings
// Common Searches: how to delete printer settings from all sheets using Aspose.Cells | remove page setup printer configuration before saving workbook .NET | Aspose.Cells clear printer settings from worksheets | C# clean Excel file printer settings Aspose | strip printer settings from Excel workbook programmatically
// Developer Intent: Load a workbook, clear printer settings on every worksheet, and save the file as a clean XLSX.
// Use Cases: Sanitize a workbook before distribution to avoid printer‑specific issues on recipient machines. | Generate a template workbook programmatically and ensure it contains no residual printer configuration. | Process user‑uploaded Excel files on a server, strip embedded printer settings, and store the cleaned version.
// AI Prompts: Generate C# code with Aspose.Cells that iterates through all worksheets and sets PageSetup.PrinterSettings to null. | Explain how to confirm that printer settings have been removed after saving a workbook with Aspose.Cells. | Provide a fallback strategy for when the source Excel file is missing, creating a new workbook before cleaning printer settings.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an existing Excel file (or creates a new workbook if missing), iterates through every worksheet, clears its printer configuration by setting PageSetup.PrinterSettings to null, and saves the result as a clean XLSX file.
    public class RemovePrinterSettingsDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Define source and output file paths
            string sourcePath = "input.xlsx";
            string outputPath = "clean.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(sourcePath))
            {
                workbook = new Workbook(sourcePath);
            }
            else
            {
                Console.WriteLine($"Source file '{sourcePath}' not found. Creating a new workbook.");
                workbook = new Workbook();
                // Ensure at least one worksheet exists
                if (workbook.Worksheets.Count == 0)
                {
                    workbook.Worksheets.Add();
                }
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Remove printer settings from each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Setting PrinterSettings to null clears stored printer configuration
                sheet.PageSetup.PrinterSettings = null;
            }

            // Save the cleaned workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Printer settings removed and workbook saved to: {outputPath}");
        }
    }
}
