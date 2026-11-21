// Title: C# – Validate Freeze Column Index Within Worksheet Range Before Using FreezePanes in Aspose.Cells
// Common Searches: C# Aspose.Cells how to verify freeze column index is inside worksheet range | Aspose.Cells FreezePanes column index out of range handling in .NET | validate column index before calling FreezePanes with Aspose.Cells | prevent error when freezing columns beyond Excel's column limit using Aspose.Cells C# | check worksheet column bounds prior to FreezePanes in Aspose.Cells
// Tags: freeze column index validation Aspose.Cells | worksheet column range check Aspose.Cells | FreezePanes column bounds C# | Excel column limit handling Aspose.Cells | C# Aspose.Cells freeze panes error prevention

using System;
using System.IO;
using Aspose.Cells;

namespace FreezeColumnExample
{
    // The example creates a workbook, determines the maximum valid column index (capped at Excel's 16383 limit), validates a zero‑based freeze column index against the worksheet's column range, and calls Worksheet.FreezePanes only when the index is within bounds. It also ensures the output directory exists before saving the workbook, providing safe handling of out‑of‑range freeze operations.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Desired freeze column index (zero‑based)
                int freezeColumnIndex = 5; // example value

                // Determine maximum column index (0‑16383)
                int maxColumnIndex = Math.Min(16383, Math.Max(worksheet.Cells.MaxColumn, 0));

                if (freezeColumnIndex >= 0 && freezeColumnIndex <= maxColumnIndex)
                {
                    // Freeze panes at the specified column.
                    // Parameters: (row, column, totalRows, totalColumns)
                    // To freeze only columns, set row to 0.
                    worksheet.FreezePanes(0, freezeColumnIndex, 0, freezeColumnIndex);
                }
                else
                {
                    Console.WriteLine("Error: Freeze column index is outside the worksheet's column range.");
                }

                // Define output path
                string outputPath = "FreezeColumnValidated.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
