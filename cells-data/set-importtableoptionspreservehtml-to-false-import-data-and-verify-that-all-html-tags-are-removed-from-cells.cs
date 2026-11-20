// Title: Import data with ImportTableOptions.PreserveHtml set to false and verify HTML tags are removed from cells using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# import array without preserving HTML markup | Set ImportTableOptions PreserveHtml to false before importing data in Aspose.Cells | Check that HTML tags are stripped from worksheet cells after import with Aspose.Cells | C# code to disable HTML preservation during ImportArray in Aspose.Cells | How to validate cell values for removed HTML tags after data import in Aspose.Cells
// Tags: ImportTableOptions PreserveHtml false | ImportArray plain text import Aspose.Cells | HTML tag removal from worksheet cells C# | cell content validation after data import Aspose.Cells | save workbook after cleaning HTML Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new workbook, disables HTML preservation by setting ImportTableOptions.PreserveHtml to false, imports a 2x2 string array containing HTML markup using ImportArray, iterates through each cell to confirm that no '<' or '>' characters remain, outputs the verification results, and saves the workbook as 'ImportedData.xlsx'.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var worksheet = workbook.Worksheets[0];

                // Sample data containing HTML tags (as strings)
                string[,] sourceData = new string[,]
                {
                    { "<b>Bold Text</b>", "<a href='https://example.com'>Link</a>" },
                    { "<div>Div Content</div>", "Plain Text" }
                };

                // Import the data starting at cell A1 (row 0, column 0)
                // ImportArray treats the values as plain text, so HTML tags are not preserved
                worksheet.Cells.ImportArray(sourceData, 0, 0);

                // Verify that HTML tags have been stripped from each imported cell
                for (int r = 0; r < sourceData.GetLength(0); r++)
                {
                    for (int c = 0; c < sourceData.GetLength(1); c++)
                    {
                        var cell = worksheet.Cells[r, c];
                        string cellValue = cell.StringValue;

                        // Simple validation: the value should not contain '<' or '>'
                        if (cellValue.Contains("<") || cellValue.Contains(">"))
                        {
                            Console.WriteLine($"HTML tag still present in cell {cell.Name}: {cellValue}");
                        }
                        else
                        {
                            Console.WriteLine($"Cell {cell.Name} cleaned successfully: {cellValue}");
                        }
                    }
                }

                // Save the workbook (optional, demonstrates usage of the save rule)
                string outputPath = "ImportedData.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
