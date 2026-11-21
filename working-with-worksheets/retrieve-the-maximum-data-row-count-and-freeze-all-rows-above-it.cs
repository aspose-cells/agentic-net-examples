// Title: How to get the last populated row and freeze rows above it with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get last data row index in worksheet | Freeze rows above last populated row using Aspose.Cells FreezePanes | Determine maximum data row and apply freeze panes in Aspose.Cells | Programmatically freeze top rows in Excel with Aspose.Cells .NET | Retrieve MaxDataRow and set FreezePanes in Aspose.Cells example | C# code to freeze rows up to the last used row in an Excel workbook
// Tags: Aspose.Cells MaxDataRow retrieval | Aspose.Cells FreezePanes usage | C# freeze rows above last data row Excel | determine last populated row Aspose.Cells | Excel worksheet row freezing with Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, uses Cells.MaxDataRow to find the zero‑based index of the last row containing data, converts it to a 1‑based row number, applies FreezePanes to lock all rows above that point, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve the maximum data row index (zero‑based)
                int maxDataRowIndex = sheet.Cells.MaxDataRow;

                // Freeze all rows above the last data row.
                // FreezePanes(row, column, totalRows, totalColumns) uses 1‑based indices.
                // Setting column and totalColumns to 0 means no columns are frozen.
                int splitRow = maxDataRowIndex + 1; // convert to 1‑based
                sheet.FreezePanes(splitRow, 0, splitRow, 0);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
