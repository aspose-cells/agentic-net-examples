// Title: C# – Rename an Excel defined name with Aspose.Cells (ReportData → SummaryData)
// Description: Loads (or creates) an Excel workbook, retrieves the defined name "ReportData", changes its Name.Text to "SummaryData", and saves the modified file. Demonstrates safe handling when the name is missing.
// Keywords: Aspose.Cells | C# | rename defined name | Name.Text | Excel named range | Workbook.Save | ReportData | SummaryData | Aspose.Cells for .NET | named range rename | Excel automation
// Common Searches: Aspose.Cells rename defined name C# | Change Name.Text property Aspose.Cells | How to rename Excel named range using Aspose.Cells .NET | C# code to rename ReportData to SummaryData | Update named range programmatically Aspose.Cells
// Developer Intent: Rename the existing defined name "ReportData" to "SummaryData" in an Excel workbook using Aspose.Cells for .NET and persist the change.
// Use Cases: Refactor legacy named ranges after a data model change. | Enforce consistent naming across a batch of generated reports. | Replace placeholder names in a template before distribution. | Automate workbook cleanup by renaming obsolete ranges.
// AI Prompts: Write C# code with Aspose.Cells that finds a defined name 'ReportData', changes its Text to 'SummaryData', and saves the workbook, handling the case where the name does not exist. | Show a C# Aspose.Cells snippet that creates a workbook, adds a dummy defined name, renames it to 'SummaryData', and writes the file to disk. | Provide a GitHub‑style README example for renaming an Excel named range using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

// Loads (or creates) an Excel workbook, retrieves the defined name "ReportData", changes its Name.Text to "SummaryData", and saves the modified file. Demonstrates safe handling when the name is missing.
class Program
{
    static void Main()
    {
        const string inputPath = "ReportDataWorkbook.xlsx";
        const string outputPath = "ModifiedReportDataWorkbook.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one with a placeholder defined name.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Add a dummy defined name "ReportData" pointing to cell A1.
                int nameIndex = workbook.Worksheets.Names.Add("ReportData");
                workbook.Worksheets.Names[nameIndex].RefersTo = $"{sheet.Name}!$A$1";
            }

            // Retrieve the defined name "ReportData".
            Name reportName = workbook.Worksheets.Names["ReportData"];

            // If the name exists, rename it to "SummaryData".
            if (reportName != null)
            {
                reportName.Text = "SummaryData";
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
