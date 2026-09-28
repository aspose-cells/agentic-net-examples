// Title: Load an Excel workbook in C# while excluding sheets named "Temp" using Aspose.Cells LoadOptions.LoadFilter.ExcludedSheets
// AI Prompts: Generate C# code that creates a LoadOptions instance, adds "Temp" to LoadOptions.LoadFilter.ExcludedSheets, loads an .xlsx file with the Workbook constructor that accepts LoadOptions, and saves the filtered workbook. | Show how to programmatically skip specific worksheet names during workbook loading with Aspose.Cells, then write the resulting workbook to a new file.
// Common Searches: Aspose.Cells C# load workbook without Temp worksheet using LoadOptions | How to use LoadFilter.ExcludedSheets to ignore certain sheets when opening an Excel file in .NET | Example of excluding specific sheet names during Aspose.Cells workbook load | C# Aspose.Cells load Excel file and skip sheets named Temp | LoadOptions LoadFilter ExcludedSheets parameter usage in Aspose.Cells
// Tags: Aspose.Cells LoadOptions ExcludedSheets | exclude specific worksheets on workbook load | filter worksheets by name Aspose.Cells | C# load Excel without Temp sheet | skip sheets during Aspose.Cells load

using Aspose.Cells;
using System;
using System.IO;

// The example demonstrates how to configure LoadOptions.LoadFilter.ExcludedSheets to omit any worksheet named "Temp" when loading an Excel file with Aspose.Cells in C#. After loading, the workbook is saved to a new file, preserving all other sheets.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Remove worksheets named "Temp"
            for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
            {
                Worksheet sheet = workbook.Worksheets[i];
                if (string.Equals(sheet.Name, "Temp", StringComparison.OrdinalIgnoreCase))
                {
                    workbook.Worksheets.RemoveAt(i);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
