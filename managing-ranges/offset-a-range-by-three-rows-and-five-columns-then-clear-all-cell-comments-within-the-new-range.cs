// Title: Shift an Excel range by 3 rows and 5 columns and clear all comments in the shifted range using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a workbook, creates a range shifted three rows down and five columns right from an existing range, and removes any cell comments within the new area using Aspose.Cells. | Write a function that accepts a source range, computes its displaced position, iterates through the displaced cells, deletes attached comments, and saves the workbook, including file existence checks and directory creation.
// Common Searches: Aspose.Cells C# offset range and remove comments | How to delete comments from a range moved down rows in Excel using Aspose | C# example for creating displaced range and clearing cell comments | Programmatically clear comments after shifting a range with Aspose.Cells | Remove worksheet comments from a range shifted by rows and columns in .NET
// Tags: range displacement Aspose.Cells | clear cell comments Aspose.Cells | Aspose.Cells workbook comment cleanup | C# Excel range manipulation | shifted range creation Aspose.Cells | Excel file comment removal .NET

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample loads an existing workbook, defines a source range (A1:C10), computes a new range offset by three rows and five columns, creates that displaced range, iterates through each cell in the new area, removes any associated comments via the worksheet's Comments collection, and saves the modified workbook to a new file while handling missing input files and ensuring the output directory exists.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Define the original range (example: A1:C10)
            AsposeRange originalRange = sheet.Cells.CreateRange("A1", "C10");

            // Calculate the offset position: 3 rows down, 5 columns to the right
            int offsetFirstRow = originalRange.FirstRow + 3;
            int offsetFirstColumn = originalRange.FirstColumn + 5;
            int rowCount = originalRange.RowCount;
            int columnCount = originalRange.ColumnCount;

            // Create the new offset range
            AsposeRange offsetRange = sheet.Cells.CreateRange(offsetFirstRow, offsetFirstColumn, rowCount, columnCount);

            // Iterate through each cell in the offset range and remove its comment if it exists
            for (int i = 0; i < offsetRange.RowCount; i++)
            {
                for (int j = 0; j < offsetRange.ColumnCount; j++)
                {
                    try
                    {
                        Cell cell = offsetRange[i, j];
                        if (cell?.Comment != null)
                        {
                            int commentIndex = sheet.Comments.IndexOf(cell.Comment);
                            if (commentIndex >= 0)
                            {
                                sheet.Comments.RemoveAt(commentIndex);
                            }
                        }
                    }
                    catch (Exception exCell)
                    {
                        Console.WriteLine($"Error processing cell at offset ({i},{j}): {exCell.Message}");
                    }
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
