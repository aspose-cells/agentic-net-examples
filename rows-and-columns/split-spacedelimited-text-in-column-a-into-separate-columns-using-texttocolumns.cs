// Title: How to split space‑delimited values in column A into separate columns with Aspose.Cells TextToColumns in C#
// Common Searches: C# Aspose.Cells TextToColumns example using space as separator | split column A values into multiple columns with TxtLoadOptions in .NET | convert numeric strings to numbers while using TextToColumns in Aspose.Cells | save workbook after TextToColumns operation to XLSX using C# | populate Excel column with space‑delimited strings and parse with Aspose.Cells
// Tags: Aspose.Cells TextToColumns with space separator | C# TxtLoadOptions for delimited data | numeric conversion during TextToColumns | export parsed data to XLSX using Aspose.Cells | populate worksheet column with delimited strings C#

using System;
using Aspose.Cells;

namespace AsposeCellsTextToColumnsDemo
{
    // The sample creates a new workbook, fills column A with space‑delimited strings, configures TxtLoadOptions to use a space character as the separator and to convert numeric data, then calls Cells.TextToColumns to split the data into separate columns, prints the results, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Sample space‑delimited data to be placed in column A
                string[] sourceData = new string[]
                {
                    "John Doe 30",
                    "Jane Smith 25",
                    "Bob Johnson 40",
                    "Alice Brown 28",
                    "Mike Davis 35"
                };

                // Populate column A (zero‑based column index 0) with the sample data
                for (int i = 0; i < sourceData.Length; i++)
                {
                    cells[i, 0].PutValue(sourceData[i]);
                }

                // Set up options for TextToColumns: use space as the separator
                TxtLoadOptions txtOptions = new TxtLoadOptions
                {
                    Separator = ' ',               // Separator expects a char, not a string
                    ConvertNumericData = true
                };

                // Apply TextToColumns using the options.
                // -1 for columnCount lets Aspose.Cells determine the required number of columns.
                cells.TextToColumns(0, 0, -1, txtOptions);

                // Verify the result by printing the rows to the console
                Console.WriteLine("Result after TextToColumns:");
                for (int row = 0; row < sourceData.Length; row++)
                {
                    string colA = cells[row, 0].StringValue;
                    string colB = cells[row, 1].StringValue;
                    string colC = cells[row, 2].StringValue;
                    Console.WriteLine($"Row {row + 1}: A=\"{colA}\", B=\"{colB}\", C=\"{colC}\"");
                }

                // Save the workbook to an XLSX file
                string outputPath = "TextToColumnsResult.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to \"{outputPath}\"");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
