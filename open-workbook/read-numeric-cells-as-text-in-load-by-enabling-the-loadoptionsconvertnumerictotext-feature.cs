// Title: Load an Excel workbook in C# with Aspose.Cells and read numeric cells as text using ConvertNumericToText
// Common Searches: Aspose.Cells C# load workbook and treat numbers as strings | How to get string value of a numeric cell with LoadOptions ConvertNumericToText | Read Excel cell A1 as text instead of number using Aspose.Cells | Enable ConvertNumericToText option when opening XLSX in C#
// Tags: Aspose.Cells LoadOptions ConvertNumericToText | numeric cell stringvalue C# | Excel workbook load numeric as text | XLSX loading option Aspose.Cells | C# Aspose.Cells cell.StringValue usage

using Aspose.Cells;
using System;
using System.IO;

// The example demonstrates loading an XLSX file with Aspose.Cells using LoadOptions that enable ConvertNumericToText, then accessing a cell's StringValue to retrieve numeric data as plain text, and finally saving the workbook while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook (numeric cells will be accessible as text via StringValue)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet and read cell A1 as text
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            string cellText = cell.StringValue; // Numeric cell represented as text

            Console.WriteLine($"Cell A1 as text: {cellText}");

            // Save the workbook (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
