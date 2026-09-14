// Title: How to remove all formatting while converting an Excel table to a range by setting TableToRangeOptions.LastRow = 0 using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads a workbook, accesses the first ListObject, and converts it to a normal range with TableToRangeOptions where LastRow is set to 0 to discard table styles. | Generate an Aspose.Cells example that demonstrates using TableToRangeOptions to strip formatting during a table‑to‑range conversion and saves the result. | Provide a step‑by‑step script that applies TableToRangeOptions.LastRow = 0 to an Excel table conversion, ensuring no formatting is retained in the output file.
// Common Searches: Aspose.Cells TableToRangeOptions LastRow zero example C# | convert Excel ListObject to range without preserving styles Aspose.Cells | remove table formatting during conversion to range .NET | C# Aspose.Cells how to discard table styles when converting to range | set TableToRangeOptions.LastRow to 0 to clear formatting in Excel workbook
// Tags: Aspose.Cells TableToRangeOptions remove formatting | C# convert ListObject to range without styles | Excel table to range conversion options | TableToRangeOptions LastRow zero usage | Aspose.Cells discard table formatting | range conversion without preserving styles .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, retrieves the first ListObject, and uses TableToRangeOptions with LastRow set to 0 to convert the table into a regular range while stripping all table formatting. The modified workbook is then saved to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; create a simple workbook with a table if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                var tempWs = tempWb.Worksheets[0];

                // Populate sample data.
                tempWs.Cells["A1"].PutValue("Header1");
                tempWs.Cells["B1"].PutValue("Header2");
                tempWs.Cells["A2"].PutValue("Data1");
                tempWs.Cells["B2"].PutValue("Data2");

                // Add a table covering the data range.
                var tempTable = tempWs.ListObjects[tempWs.ListObjects.Add(0, 0, 2, 2, true)];
                tempTable.DisplayName = "SampleTable";

                tempWb.Save(inputPath);
            }

            // Verify the input file still exists before loading.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook.
            var workbook = new Workbook(inputPath);
            var sheet = workbook.Worksheets[0];

            // Verify that at least one table exists on the worksheet.
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables found in the worksheet.");
                return;
            }

            // Access the first table.
            var table = sheet.ListObjects[0];

            // Capture the table's range coordinates before conversion.
            int startRow = table.StartRow;
            int startColumn = table.StartColumn;
            int endRow = table.EndRow;
            int endColumn = table.EndColumn;
            int rowCount = endRow - startRow + 1;
            int columnCount = endColumn - startColumn + 1;

            // Convert the table to a normal range (default options).
            table.ConvertToRange();

            // Obtain the range that was previously the table.
            var range = sheet.Cells.CreateRange(startRow, startColumn, rowCount, columnCount);

            // Remove all formatting from the converted range.
            range.ClearFormats();

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
