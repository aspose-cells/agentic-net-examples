// Title: Copy a cell range and transpose rows to columns in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that copies a specified range and writes it transposed to a target cell, preserving values and formatting. | Show how to implement a reusable method in Aspose.Cells for .NET that copies any range and flips rows and columns while keeping cell styles. | Generate an example that loads a workbook, copies range A1:C3, transposes it to cell E1, and saves the file using Aspose.Cells.
// Common Searches: Aspose.Cells C# copy range and transpose to another location | how to preserve cell formatting when transposing a range with Aspose.Cells | transpose rows to columns while copying range in Excel using Aspose.Cells for .NET | C# Aspose.Cells example for copying range A1:C3 to E1 with transposition | manual range copy with style preservation Aspose.Cells .NET
// Tags: range copy with transposition Aspose.Cells | preserve cell formatting during range copy .NET | transpose rows to columns Excel Aspose.Cells | copy range A1:C3 to E1 Aspose.Cells example | save workbook after transposed range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Copies the range A1:C3 from the first worksheet, transposes it to start at cell E1 while preserving values and cell styles, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the source worksheet (first sheet in this example)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Define the source range to copy (e.g., A1:C3)
            AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:C3");

            // Destination start cell where the transposed data will be placed (e.g., E1)
            int destRow = 0;      // Row index for row 1
            int destColumn = 4;   // Column index for column E

            // Manually copy with transposition (preserving values and styles)
            int rowCount = sourceRange.RowCount;
            int columnCount = sourceRange.ColumnCount;

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    // Source cell
                    Cell srcCell = sourceSheet.Cells[sourceRange.FirstRow + i, sourceRange.FirstColumn + j];

                    // Destination cell (transposed)
                    Cell destCell = sourceSheet.Cells[destRow + j, destColumn + i];

                    // Copy value
                    destCell.PutValue(srcCell.Value);

                    // Copy style
                    destCell.SetStyle(srcCell.GetStyle());
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

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
