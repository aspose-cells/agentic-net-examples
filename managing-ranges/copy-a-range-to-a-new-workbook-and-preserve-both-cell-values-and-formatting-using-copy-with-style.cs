// Title: Copy a cell range with values, formulas, and formatting to a new workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that copies the A1:C10 range from an existing workbook to a newly created workbook while preserving values, formulas, and cell styles with Aspose.Cells. | Show how to use Aspose.Cells PasteOptions with PasteType.All to duplicate a source range into a destination range in another workbook, keeping all formatting. | Write a C# program that creates a sample source.xlsx if missing, defines a range, and copies it to dest.xlsx with full formatting retention using Aspose.Cells.
// Common Searches: Aspose.Cells C# copy range to another workbook preserving formatting | How to use PasteOptions PasteType.All for range copy in Aspose.Cells | Copy Excel cells with styles to a new file using Aspose.Cells .NET | Example of copying A1:C10 from one workbook to another with Aspose.Cells | Preserve formulas and cell formatting when moving a range between workbooks in C#
// Tags: copy range with formatting Aspose.Cells | PasteOptions PasteType.All C# | duplicate Excel range to new workbook .NET | preserve cell styles Aspose.Cells | range copy between workbooks example

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // The program checks for source.xlsx, creates sample data if it doesn't exist, loads the workbook, defines the A1:C10 range, creates a new workbook, and copies the range to the destination worksheet using PasteOptions with PasteType.All to retain values, formulas, and formatting, then saves dest.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            const string sourcePath = "source.xlsx";
            const string destPath = "dest.xlsx";

            try
            {
                // Ensure the source file exists; create a simple workbook if it does not.
                if (!File.Exists(sourcePath))
                {
                    var tempWorkbook = new Workbook();
                    var tempSheet = tempWorkbook.Worksheets[0];
                    // Populate some sample data in A1:C10
                    for (int row = 0; row < 10; row++)
                    {
                        for (int col = 0; col < 3; col++)
                        {
                            tempSheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                        }
                    }
                    tempWorkbook.Save(sourcePath);
                }

                // Load the source workbook
                Workbook srcWorkbook = new Workbook(sourcePath);
                Worksheet srcWorksheet = srcWorkbook.Worksheets[0];

                // Define the range to copy
                AsposeRange srcRange = srcWorksheet.Cells.CreateRange("A1:C10");

                // Create a new workbook for the destination
                Workbook destWorkbook = new Workbook();
                Worksheet destWorksheet = destWorkbook.Worksheets[0];

                // Define the target range in the destination worksheet
                AsposeRange destRange = destWorksheet.Cells.CreateRange("A1:C10");

                // Copy the source range to the destination range, preserving values and formatting
                var pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All, // copy values, formats, formulas, etc.
                    SkipBlanks = false,
                    Transpose = false
                };
                srcRange.Copy(destRange, pasteOptions);

                // Save the destination workbook
                destWorkbook.Save(destPath);

                Console.WriteLine($"Range copied successfully from '{sourcePath}' to '{destPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
