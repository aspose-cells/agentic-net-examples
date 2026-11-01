// Title: Apply customizable freeze panes to lock header rows and columns in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write a C# method that opens an Excel workbook, freezes the top N rows and left M columns with Aspose.Cells, and saves the modified file. | Create a console program that reads input file path, output file path, freeze row index, and freeze column index from command‑line arguments and applies FreezePanes to the first worksheet using Aspose.Cells. | Implement robust error handling that checks for the source file, creates missing output directories, and wraps exceptions when applying FreezePanes in Aspose.Cells.
// Common Searches: Aspose.Cells C# freeze panes based on row and column indices | how to lock header rows and columns in Excel using Aspose.Cells .NET | C# console app to apply FreezePanes to the first worksheet with Aspose.Cells | example of saving an Excel file after setting FreezePanes with Aspose.Cells | validate input file and create output folder when using Aspose.Cells FreezePanes
// Tags: Aspose.Cells FreezePanes C# | custom header lock Excel Aspose.Cells | load and save workbook Aspose.Cells .NET | command line Excel processing Aspose.Cells | error handling Aspose.Cells file operations

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example validates the source file, loads it into a Workbook, accesses the first worksheet, applies FreezePanes using user‑provided row and column indices, ensures the output directory exists, and saves the updated workbook with comprehensive exception handling.
    public class FreezePaneHelper
    {
        /// <param name="inputFilePath">Path to the source Excel file.</param>
        /// <param name="outputFilePath">Path where the modified file will be saved.</param>
        /// <param name="freezeRow">Number of rows to freeze (rows above this index will be locked).</param>
        /// <param name="freezeColumn">Number of columns to freeze (columns left of this index will be locked).</param>
        public void ApplyFreezePanes(string inputFilePath, string outputFilePath, int freezeRow, int freezeColumn)
        {
            // Validate input file existence
            if (!File.Exists(inputFilePath))
                throw new FileNotFoundException($"Input file not found: {inputFilePath}");

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputFilePath);

                // Access the first worksheet (modify index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Apply freeze panes: rows above freezeRow and columns left of freezeColumn
                worksheet.FreezePanes(freezeRow, freezeColumn, freezeRow, freezeColumn);

                // Ensure the output directory exists
                string outDir = Path.GetDirectoryName(outputFilePath);
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                    Directory.CreateDirectory(outDir);

                // Save the modified workbook
                workbook.Save(outputFilePath);
            }
            catch (Exception ex)
            {
                // Wrap exception for clearer context
                throw new ApplicationException("Error applying freeze panes.", ex);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // args[0] = input file path, args[1] = output file path, args[2] = freeze row, args[3] = freeze column
            if (args.Length < 4)
            {
                Console.WriteLine("Usage: <inputFilePath> <outputFilePath> <freezeRow> <freezeColumn>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            if (!int.TryParse(args[2], out int freezeRow))
            {
                Console.WriteLine("Invalid freezeRow value.");
                return;
            }

            if (!int.TryParse(args[3], out int freezeColumn))
            {
                Console.WriteLine("Invalid freezeColumn value.");
                return;
            }

            try
            {
                var helper = new FreezePaneHelper();
                helper.ApplyFreezePanes(inputPath, outputPath, freezeRow, freezeColumn);
                Console.WriteLine($"Freeze panes applied successfully. Output saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
