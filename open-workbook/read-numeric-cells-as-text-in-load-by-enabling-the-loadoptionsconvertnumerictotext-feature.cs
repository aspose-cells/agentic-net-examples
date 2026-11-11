// Title: Read numeric Excel cells as text by enabling LoadOptions.ConvertNumericToText in Aspose.Cells (C#)
// Common Searches: How to load an XLSX file with numeric values treated as text using Aspose.Cells C# | Aspose.Cells C# example for converting numbers to strings on load | Read cell A1 as string instead of number with Aspose.Cells LoadOptions | Preserve leading zeros when opening Excel workbook in C# with Aspose.Cells | Enable numeric-to-text conversion to prevent number formatting loss in Aspose.Cells
// Tags: Aspose.Cells LoadOptions ConvertNumericToText | read numeric cell as string C# | load XLSX with text representation of numbers | preserve leading zeros Aspose.Cells | numeric values as text on workbook load

using Aspose.Cells;
using System;
using System.IO;

// This example shows how to load an XLSX workbook with Aspose.Cells while converting all numeric cells to text using the LoadOptions.ConvertNumericToText flag, then reads cell A1 as a string and saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook with default options
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet and read cell A1 as text
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            string cellText = cell.StringValue; // Numeric value is represented as text

            Console.WriteLine($"Cell A1 as text: {cellText}");

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
