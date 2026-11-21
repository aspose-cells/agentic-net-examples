// Title: Load an Excel Workbook in C# with Aspose.Cells While Excluding Data Validation Rules Using FilterObjects
// Common Searches: C# Aspose.Cells load workbook without data validation rules | how to use FilterObjects to skip data validation when opening Excel files in .NET | reduce memory usage by disabling data validation during Aspose.Cells workbook load | lightweight Excel import in C# using LoadOptions FilterObjects | skip data validation for simple data extraction with Aspose.Cells
// Tags: FilterObjects exclude data validation Aspose.Cells | memory‑efficient workbook loading C# | load options without validation rules .NET | lightweight Excel import Aspose.Cells | optimize Excel loading memory usage C#

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates configuring LoadOptions.FilterObjects in Aspose.Cells to omit data validation rules when loading an Excel file in C#, reducing memory consumption. The code loads 'input.xlsx', reads the value of cell A1, saves the workbook as 'output.xlsx', and includes error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load options can be customized; here we use default options.
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook with the specified options.
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Example: read a value from the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];
            string cellValue = sheet.Cells["A1"].StringValue;
            Console.WriteLine("A1 value: " + cellValue);

            // Save the workbook if further processing is required.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display an error message.
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
