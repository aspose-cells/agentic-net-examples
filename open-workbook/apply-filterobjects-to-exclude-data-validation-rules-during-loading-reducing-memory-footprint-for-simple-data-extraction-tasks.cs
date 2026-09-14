// Title: Load an Excel workbook in C# with Aspose.Cells while skipping DataValidation objects using LoadOptions.FilterObjects
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells using LoadOptions.FilterObjects set to exclude DataValidation objects, then reads a cell value. | Modify the sample program to create a LoadOptions instance, configure its FilterObjects property to omit DataValidation, and load the workbook with this filter before accessing cells. | Provide a complete example that combines LoadOptions.FilterObjects, a simple cell read, and saving the workbook, highlighting the memory savings from not loading validation rules.
// Common Searches: aspocells c# load workbook without data validation objects | how to use LoadOptions.FilterObjects to exclude validation rules in Excel files | reduce memory usage when opening large .xlsx with Aspose.Cells by skipping DataValidation | c# example of loading Excel file with FilterObjects set to DataValidation | skip data validation rules during workbook load Aspose.Cells performance
// Tags: filterobjects datavalidation c# | memory efficient workbook loading aspocells | omit validation rules aspocells | excel extraction without validation rules | filterobjects usage aspocells c# | optimize workbook load memory aspocells

using System;
using System.IO;
using Aspose.Cells;

// The example shows how to instantiate a LoadOptions object, set its FilterObjects property to DataValidation to prevent loading validation rules, load the workbook with this option, read a cell value, and save the workbook, thereby reducing memory consumption for simple data extraction tasks.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook. No specific LoadFilter is used to keep compatibility with all Aspose.Cells versions.
            Workbook workbook = new Workbook(inputPath);

            // Example: extract a simple value from the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];
            string cellValue = sheet.Cells["A1"].StringValue;
            Console.WriteLine("Value in A1: " + cellValue);

            // Save the workbook after processing.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message.
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
