// Title: How to apply the '#,##0.00' numeric format to a specific cell range in an Excel file with Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a Style with the custom numeric format '#,##0.00' and applies it to cells B1:B10 using Aspose.Cells. | Show how to use StyleFlag to limit formatting to the number format only, then save the workbook after applying the currency style.
// Common Searches: Aspose.Cells C# set number format '#,##0.00' for a range of cells | How to format Excel cells as currency using Aspose.Cells .NET | Apply custom numeric format to B1:B10 with Aspose.Cells library | Using StyleFlag to change only number format in Aspose.Cells workbook
// Tags: numeric format styling Aspose.Cells | StyleFlag number format application | currency display formatting Excel | cell range formatting Aspose.Cells | C# workbook number format customization

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program loads an existing workbook, creates a Style with the custom numeric format '#,##0.00', applies it to cells B1:B10 via a StyleFlag that targets only the number format, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException($"The input file '{inputPath}' was not found.");
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (or specify the desired one)
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a new style object for currency formatting
                Style currencyStyle = workbook.CreateStyle();
                currencyStyle.Custom = "#,##0.00";

                // Define which style attributes to apply (only the number format)
                StyleFlag styleFlag = new StyleFlag
                {
                    NumberFormat = true
                };

                // Define the target range B1:B10 (row 0‑9, column 1)
                Aspose.Cells.Range range = worksheet.Cells.CreateRange(0, 1, 10, 1);
                range.ApplyStyle(currencyStyle, styleFlag);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
