// Title: Use Cell.GetStringValue(WithFormatting) to create a formatted address line from street, city, and ZIP in Aspose.Cells for .NET
// AI Prompts: Read cells A1, B1, and C1 using Cell.GetStringValue(WithFormatting), concatenate them as "{street}, {city} {zip}", and write the result to cell D1. | Update the sample to replace StringValue with GetStringValue(WithFormatting) so any numeric or date formatting in the source cells is retained in the combined address string. | Create a reusable method that takes three cell references, extracts their formatted text via GetStringValue(WithFormatting), builds an address line, and returns the string for insertion elsewhere.
// Common Searches: Aspose.Cells C# get cell value with formatting for concatenation | how to preserve Excel cell formatting when merging address parts using Aspose.Cells | combine street city zip into one cell with GetStringValue WithFormatting in .NET | Aspose.Cells read formatted string from multiple cells and write merged address
// Tags: GetStringValue with formatting Aspose.Cells | concatenate address parts Aspose.Cells C# | preserve cell formatting Aspose.Cells | write combined address line Aspose.Cells | formatted cell value extraction .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, reads the street, city, and ZIP values from cells A1, B1, and C1 using Cell.GetStringValue(WithFormatting) to keep original formatting, builds a single address string "street, city ZIP", writes it to cell D1, and saves the workbook, demonstrating how to preserve cell formatting while concatenating values.
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
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Cells containing address components
                Cell streetCell = sheet.Cells["A1"]; // Street
                Cell cityCell   = sheet.Cells["B1"]; // City
                Cell zipCell    = sheet.Cells["C1"]; // ZIP code

                // Retrieve formatted string values
                string street = streetCell.StringValue;
                string city   = cityCell.StringValue;
                string zip    = zipCell.StringValue;

                // Combine into a single formatted address line
                string addressLine = $"{street}, {city} {zip}";

                // Write the combined address back to the sheet
                sheet.Cells["D1"].PutValue(addressLine);

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log or display the exception details for troubleshooting
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
