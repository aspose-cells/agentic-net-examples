// Title: Use Aspose.Cells TextToColumns in C# to split semicolon‑delimited values in a worksheet column
// AI Prompts: Generate C# code that opens an existing Excel file, configures TxtLoadOptions with a ';' separator, applies TextToColumns on column A, and saves the result to a new workbook. | Write a C# program that verifies the presence of input.xlsx, creates the output folder if it does not exist, and splits semicolon‑separated strings in the first column into separate columns using Aspose.Cells. | Provide a C# example that includes try‑catch error handling while loading a workbook, applying TextToColumns with a custom delimiter, and writing the transformed data to output.xlsx.
// Common Searches: Aspose.Cells C# split column values by semicolon using TextToColumns | How to parse semicolon separated data in Excel with Aspose.Cells in .NET | C# example for TextToColumns with custom delimiter in Aspose.Cells | Convert a single column of ; delimited strings to multiple columns using Aspose.Cells
// Tags: Aspose.Cells TextToColumns custom separator | C# split semicolon delimited Excel column | TxtLoadOptions separator property usage | Load and save workbook with Aspose.Cells | Handle missing input file in C# Excel processing

using System;
using System.IO;
using Aspose.Cells;

// Loads input.xlsx, uses Aspose.Cells TxtLoadOptions with a ';' separator to apply TextToColumns on column A, creates the output directory if needed, and saves the transformed workbook to output.xlsx with proper error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the last used row in column A (zero‑based index)
            int lastRow = sheet.Cells.GetLastDataRow(0);
            if (lastRow < 0)
            {
                Console.WriteLine("No data found in column A.");
                return;
            }

            // Prepare TextToColumns options with semicolon delimiter
            TxtLoadOptions txtOptions = new TxtLoadOptions
            {
                Separator = ';'
            };

            // Apply TextToColumns on column A
            // totalRows = lastRow + 1 (since rows are zero‑based)
            sheet.Cells.TextToColumns(0, lastRow + 1, 0, txtOptions);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
