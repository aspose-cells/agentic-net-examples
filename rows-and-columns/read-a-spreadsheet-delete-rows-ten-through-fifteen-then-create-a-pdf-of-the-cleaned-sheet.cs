// Title: C# – Delete rows 10‑15 from an Excel worksheet and save as PDF using Aspose.Cells
// Description: Loads an .xlsx file (creates a sample workbook if missing), removes rows 10‑15 from the first worksheet via Cells.DeleteRows, ensures the output directory exists, and saves the result as a PDF with SaveFormat.Pdf. Includes basic error handling.
// Keywords: Aspose.Cells | C# | DeleteRows | Excel to PDF | remove rows | SaveFormat.Pdf | .NET | worksheet export | row deletion | PDF conversion
// Common Searches: Aspose.Cells delete rows 10 to 15 C# | Convert Excel to PDF after removing rows Aspose | How to delete specific rows in .NET Excel library | Export cleaned worksheet to PDF using Aspose.Cells | C# code sample delete rows and save as PDF
// Developer Intent: Remove a fixed range of rows from an Excel sheet and generate a PDF of the cleaned workbook.
// Use Cases: Prepare client‑ready reports by stripping placeholder rows before PDF export. | Automate data cleanup for incoming spreadsheets and deliver them as PDFs. | Create invoice PDFs after programmatically eliminating empty or obsolete rows. | Generate printable dashboards where certain rows must be omitted.
// AI Prompts: Generate C# code that uses Aspose.Cells to delete rows 10‑15 from the first worksheet and export the workbook to PDF, with handling for missing input files. | Explain step‑by‑step how to delete a row range and then save the worksheet as PDF using Aspose.Cells, including folder creation. | Modify the sample to delete rows based on a condition (e.g., empty cells) before converting to PDF. | Provide a PowerShell script that calls a compiled .NET assembly to perform the same row deletion and PDF export.

using System;
using System.IO;
using Aspose.Cells;

// Loads an .xlsx file (creates a sample workbook if missing), removes rows 10‑15 from the first worksheet via Cells.DeleteRows, ensures the output directory exists, and saves the result as a PDF with SaveFormat.Pdf. Includes basic error handling.
class DeleteRowsAndExportPdf
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = @"C:\Path\To\InputWorkbook.xlsx";
        string outputPath = @"C:\Path\To\CleanedWorkbook.pdf";

        Workbook workbook = null;

        try
        {
            // Ensure the input file exists; create a sample workbook if it does not
            if (!File.Exists(inputPath))
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                Cells wsCells = ws.Cells;

                // Populate sample data (20 rows)
                for (int i = 0; i < 20; i++)
                {
                    wsCells[i, 0].PutValue($"Row {i + 1}");
                }

                // Save the sample workbook for future runs
                workbook.Save(inputPath);
            }
            else
            {
                // Load the existing workbook
                workbook = new Workbook(inputPath);
            }

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells worksheetCells = worksheet.Cells;

            // Delete rows 10 through 15 (zero‑based index: 9, count: 6)
            worksheetCells.DeleteRows(9, 6);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
