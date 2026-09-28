// Title: How to split semicolon‑delimited text in a worksheet column using Aspose.Cells TextToColumns in C#
// AI Prompts: Generate C# code that creates a workbook, fills a column with semicolon‑separated strings, sets TxtLoadOptions.Separator to ';', and calls Cells.TextToColumns to distribute the values across separate columns. | Write a C# snippet that calculates the number of rows containing data, passes that count to the TextToColumns method, and prints each resulting cell value to the console for verification. | Provide C# code that saves the workbook after the semicolon split, showing how to specify the output file name and format with Aspose.Cells.
// Common Searches: Aspose.Cells C# split column by custom delimiter semicolon | TxtLoadOptions Separator property parse semicolon separated values in Excel | determine row count for TextToColumns method Aspose.Cells | save workbook after performing TextToColumns operation with Aspose.Cells C#
// Tags: custom separator TextToColumns Aspose.Cells | semicolon delimiter TxtLoadOptions C# | split worksheet column Aspose.Cells | retrieve max data row Aspose.Cells | export workbook after column split Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, populates column A with semicolon‑delimited strings, configures TxtLoadOptions to use ';' as the separator, invokes Cells.TextToColumns to split the data into separate columns, prints the split values, and saves the result as SemicolonSplitOutput.xlsx.
    public class TextToColumnsSemicolonDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate column A with semicolon‑separated strings
            cells["A1"].PutValue("John;Doe;30");
            cells["A2"].PutValue("Jane;Smith;28");
            cells["A3"].PutValue("Bob;Brown;45");

            // Determine how many rows contain data (zero‑based index + 1)
            int totalRows = cells.MaxDataRow + 1;

            // Set up load options to use semicolon as the delimiter
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                Separator = ';'
            };

            // Split the text in column A into separate columns
            // Parameters: start row, start column, number of rows, load options
            int createdColumns = cells.TextToColumns(0, 0, totalRows, loadOptions);

            // Display the split result in the console (optional verification)
            Console.WriteLine($"TextToColumns created {createdColumns} columns.");
            for (int r = 0; r < totalRows; r++)
            {
                for (int c = 0; c < createdColumns; c++)
                {
                    Console.Write(cells[r, c].StringValue + "\t");
                }
                Console.WriteLine();
            }

            // Save the workbook with the split data
            string outputPath = "SemicolonSplitOutput.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
