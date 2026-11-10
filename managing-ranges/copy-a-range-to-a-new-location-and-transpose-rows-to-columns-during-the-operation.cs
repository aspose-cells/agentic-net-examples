// Title: Copy an Excel range and transpose rows to columns with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells copy range and transpose to another worksheet | How to move a block of cells and flip rows and columns using Aspose.Cells | Transpose data while copying a range in Aspose.Cells for .NET | Copy range A1:C4 to E1 with rows‑to‑columns transposition in C# | Aspose.Cells example for transposing a source range to a destination range
// Tags: range copy with transposition Aspose.Cells | C# transposed range write Aspose.Cells | destination cell offset copy Aspose.Cells | extract range values as 2d array Aspose.Cells | save modified workbook Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates or loads an Excel file, defines a source range (A1:C4), extracts its values into a 2‑D array, writes those values transposed to a new location starting at E1, and saves the workbook as output.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; if not, create a simple workbook for demonstration.
                if (!File.Exists(inputPath))
                {
                    var wb = new Workbook();
                    var ws = wb.Worksheets[0];
                    ws.Cells["A1"].PutValue("R1C1");
                    ws.Cells["B1"].PutValue("R1C2");
                    ws.Cells["C1"].PutValue("R1C3");
                    ws.Cells["A2"].PutValue("R2C1");
                    ws.Cells["B2"].PutValue("R2C2");
                    ws.Cells["C2"].PutValue("R2C3");
                    ws.Cells["A3"].PutValue("R3C1");
                    ws.Cells["B3"].PutValue("R3C2");
                    ws.Cells["C3"].PutValue("R3C3");
                    ws.Cells["A4"].PutValue("R4C1");
                    ws.Cells["B4"].PutValue("R4C2");
                    ws.Cells["C4"].PutValue("R4C3");
                    wb.Save(inputPath);
                }

                // Load the existing workbook
                var workbook = new Workbook(inputPath);

                // Access the first worksheet
                var sheet = workbook.Worksheets[0];

                // Define the source range to copy (example: A1:C4)
                var srcRange = sheet.Cells.CreateRange("A1", "C4");

                // Retrieve the values of the source range as a 2‑D array
                var srcValues = srcRange.Value as object[,];
                if (srcValues == null)
                {
                    Console.WriteLine("Source range contains no data.");
                    return;
                }

                // Determine the size of the source range
                int srcRows = srcValues.GetLength(0);
                int srcCols = srcValues.GetLength(1);

                // Set the top‑left cell of the destination area (example: E1)
                int destStartRow = 0;   // Row index for E1 (zero‑based)
                int destStartCol = 4;   // Column index for E1 (zero‑based)

                // Write the transposed data to the destination range
                for (int i = 0; i < srcRows; i++)
                {
                    for (int j = 0; j < srcCols; j++)
                    {
                        // Transpose: source row i becomes destination column, source column j becomes destination row
                        sheet.Cells[destStartRow + j, destStartCol + i].PutValue(srcValues[i, j]);
                    }
                }

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
