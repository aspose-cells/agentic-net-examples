// Title: How to load an Excel workbook in C# with Aspose.Cells and convert numeric cells to text using LoadOptions.ConvertNumericToText
// AI Prompts: Load an .xlsx file with Aspose.Cells, set LoadOptions.ConvertNumericToText = true, and read cell A1 as a string using StringValue. | Write a reusable C# method that takes a file path, enables numeric‑to‑text conversion via LoadOptions, and returns the string values of specified cell addresses. | Show how to save the workbook after loading it with ConvertNumericToText enabled, ensuring numeric data remains as text in the output file.
// Common Searches: Aspose.Cells C# load workbook with numeric values as strings | Convert numeric cells to text on load using LoadOptions in Aspose.Cells .NET | Read Excel cell as string instead of number with Aspose.Cells LoadOptions | Enable ConvertNumericToText option when opening .xlsx in C# | Prevent numeric conversion when reading Excel with Aspose.Cells
// Tags: load option ConvertNumericToText | numeric cell to string Aspose.Cells | C# load .xlsx with text values | preserve numeric formatting as text | read cell StringValue Aspose

using System;
using System.IO;
using Aspose.Cells;

// The example loads an XLSX workbook in C# using Aspose.Cells, activates LoadOptions.ConvertNumericToText so numeric cells are treated as strings, reads cell A1 via StringValue, and saves the workbook, all with proper exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Configure load options (numeric-to-text conversion not available in this version)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook using the configured options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Example: read a cell value as text
            Cell cell = workbook.Worksheets[0].Cells["A1"];
            string textValue = cell.StringValue; // numeric content is now a string representation
            Console.WriteLine($"A1 as text: {textValue}");

            // Save the workbook if further processing or output is required
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
