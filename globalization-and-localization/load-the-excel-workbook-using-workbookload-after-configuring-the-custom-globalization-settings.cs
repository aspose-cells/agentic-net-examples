// Title: Load an Excel workbook with French (fr-FR) culture using Aspose.Cells LoadOptions in C#
// Common Searches: Aspose.Cells C# load workbook with French locale using LoadOptions | How to set CultureInfo to fr-FR when opening an Excel file with Aspose.Cells | C# example for loading Excel file with custom globalization settings in Aspose.Cells | LoadOptions CultureInfo parameter usage in Aspose.Cells for localized data | Reading cell values from a French‑culture workbook with Aspose.Cells in .NET
// Tags: LoadOptions CultureInfo setup | Aspose.Cells French locale loading | C# Excel workbook load with custom globalization | handling file not found in Aspose.Cells | reading cell A1 after localized load

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The sample verifies the Excel file exists, configures LoadOptions with a French (fr-FR) CultureInfo, loads the workbook using these options, prints the value of cell A1 from the first worksheet, and gracefully handles any exceptions that may occur.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string inputPath = @"C:\Data\Sample.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Configure load options with custom globalization settings (e.g., French culture)
            LoadOptions loadOptions = new LoadOptions
            {
                CultureInfo = new CultureInfo("fr-FR")
            };

            // Load the workbook using the specified load options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Example: display the value of cell A1 after loading
            Console.WriteLine("A1 value: " + workbook.Worksheets[0].Cells["A1"].StringValue);
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
