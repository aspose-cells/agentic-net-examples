// Title: Add an average subtotal row for column F at the top of an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# insert average row above data in Excel | How to calculate column F average and place it as first data row with Aspose.Cells | C# code to add a summary row at the top of a worksheet using Aspose.Cells | Insert a subtotal row after header in Excel via Aspose.Cells .NET | Aspose.Cells formula AVERAGE for column F with inserted row
// Tags: average subtotal row insertion Aspose.Cells | column F average formula Aspose.Cells | worksheet.InsertRows usage C# | add summary label to new row Aspose.Cells | Excel data range average calculation .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSubtotalExample
{
    // The example loads an existing workbook, inserts a new row directly after the header, writes an AVERAGE formula for column F covering all data rows, adds an "Average" label in column A of that row, and saves the modified file.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the data range (assuming first row is a header)
                int headerRow = 0;                         // zero‑based index
                int lastDataRow = sheet.Cells.MaxDataRow;  // zero‑based index

                // Insert a new row after the header to hold the average (subtotal)
                int subtotalRowIndex = headerRow + 1;
                sheet.Cells.InsertRows(subtotalRowIndex, 1);

                // Build the AVERAGE formula for column F (zero‑based index 5)
                // After inserting a row, data rows shift down by one
                int firstDataRow = subtotalRowIndex + 1;
                int lastDataRowAfterInsert = lastDataRow + 1;
                string formula = $"AVERAGE(F{firstDataRow + 1}:F{lastDataRowAfterInsert + 1})";

                // Place the formula in column F of the subtotal row
                Cell subtotalCell = sheet.Cells[subtotalRowIndex, 5];
                subtotalCell.Formula = formula;

                // Optional: add a label in column A
                sheet.Cells[subtotalRowIndex, 0].PutValue("Average");

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
