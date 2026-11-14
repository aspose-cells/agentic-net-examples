// Title: How to apply freeze panes at a specific row and column in an Excel file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# freeze panes after a given row and column programmatically | set freeze panes dynamically with user‑provided indices in a .NET Excel workbook | C# example to lock top rows and left columns using Worksheet.FreezePanes | load existing workbook or create a new one then apply freeze panes with Aspose.Cells | using Worksheet.FreezePanes overload with zero‑based row and column values in C#
// Tags: Aspose.Cells Worksheet.FreezePanes API | freeze panes based on row and column indices .NET | programmatic Excel freeze panes customization | load or create workbook before applying freeze panes | C# Excel worksheet freeze panes implementation

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example demonstrates a C# method that loads an existing workbook (or creates a new one), populates sample data, converts 1‑based row and column parameters to zero‑based indices, and calls Worksheet.FreezePanes to lock rows above and columns to the left of the specified cell. It also ensures the output directory exists before saving the modified workbook.
    public class FreezePanesDemo
    {
        /// <param name="freezeRow">1‑based index of the first row that should remain scrollable (rows above are frozen).</param>
        /// <param name="freezeColumn">1‑based index of the first column that should remain scrollable (columns to the left are frozen).</param>
        /// <param name="inputPath">Path to an existing workbook to load. If null or empty, a new workbook is created.</param>
        /// <param name="outputPath">Path where the modified workbook will be saved.</param>
        public void ApplyFreezePanes(int freezeRow, int freezeColumn, string inputPath, string outputPath)
        {
            try
            {
                Workbook workbook;

                // Load existing workbook if the file exists; otherwise create a new workbook.
                if (!string.IsNullOrEmpty(inputPath) && File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                Worksheet sheet = workbook.Worksheets[0];

                // Optional: populate sample data for demonstration.
                for (int i = 0; i < 20; i++)
                {
                    for (int j = 0; j < 10; j++)
                    {
                        sheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
                    }
                }

                // Convert to zero‑based indices.
                int rowIndex = Math.Max(freezeRow - 1, 0);
                int colIndex = Math.Max(freezeColumn - 1, 0);

                // Apply freeze panes using the 4‑parameter overload.
                sheet.FreezePanes(rowIndex, colIndex, rowIndex, colIndex);

                // Ensure the output directory exists.
                string outDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                {
                    Directory.CreateDirectory(outDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error applying freeze panes: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            var demo = new FreezePanesDemo();

            // Example parameters: freeze after row 5, column 3, no input file, output to Demo.xlsx.
            int freezeRow = 5;
            int freezeColumn = 3;
            string inputPath = ""; // Empty means create a new workbook.
            string outputPath = "Demo.xlsx";

            demo.ApplyFreezePanes(freezeRow, freezeColumn, inputPath, outputPath);
        }
    }
}
