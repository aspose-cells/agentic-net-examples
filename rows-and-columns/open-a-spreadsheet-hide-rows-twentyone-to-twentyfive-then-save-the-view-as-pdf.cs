// Title: Hide rows 21‑25 in an Excel worksheet and export the visible view to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# hide rows 21 to 25 before PDF conversion | How to exclude specific rows from Excel PDF export with Aspose.Cells | C# code to set row IsHidden property in Aspose.Cells workbook | Export Excel to PDF while keeping hidden rows using Aspose.Cells .NET | Programmatically hide a range of rows in Aspose.Cells and save as PDF
// Tags: set row IsHidden Aspose.Cells | hide multiple rows Aspose.Cells C# | export workbook to PDF Aspose.Cells | preserve hidden rows during PDF conversion Aspose.Cells | row visibility control Aspose.Cells .NET | convert Excel to PDF with hidden rows Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file (input.xlsx) with Aspose.Cells, iterates over rows 21‑25 (zero‑based indices 20‑24) and sets each row's IsHidden property to true when the row exists, then saves the workbook as a PDF (output.pdf). It demonstrates row‑visibility manipulation and PDF export while handling missing files and runtime exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputFile = "input.xlsx";
                const string outputFile = "output.pdf";

                // Verify that the input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputFile);

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Hide rows 21 to 25 (zero‑based indices 20‑24)
                for (int rowIndex = 20; rowIndex <= 24; rowIndex++)
                {
                    // Ensure the row index is within the existing rows collection
                    if (rowIndex < sheet.Cells.Rows.Count)
                    {
                        // Use the correct property to hide a row
                        sheet.Cells.Rows[rowIndex].IsHidden = true;
                    }
                }

                // Save the workbook as a PDF file
                workbook.Save(outputFile, SaveFormat.Pdf);
                Console.WriteLine($"PDF saved successfully to {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
