// Title: Convert string cells to numeric values in the A1:C10 range with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells convert string cells to numbers in a specific Excel range | How to change cell type from text to numeric for A1:C10 using Aspose.Cells | Count numeric cells after converting string values with Aspose.Cells in .NET | Parse numeric strings and replace with double values in Excel via Aspose.Cells C# | Save workbook after converting string values to numbers using Aspose.Cells
// Tags: Aspose.Cells string-to-number conversion | Excel range cell type transformation C# | numeric parsing of cell strings Aspose.Cells | cell value type change Aspose.Cells | save modified workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The C# program loads an Excel file (creating a simple one if missing), defines the A1:C10 range, iterates each cell, and uses double.TryParse to replace string values that represent numbers with actual numeric values. It then counts how many cells are numeric after the conversion, outputs the count, and saves the updated workbook as output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Ensure the input file exists; create a simple workbook if it does not.
        if (!File.Exists(inputPath))
        {
            try
            {
                var wbTemplate = new Workbook();
                var wsTemplate = wbTemplate.Worksheets[0];
                wsTemplate.Cells["A1"].PutValue("123.45");
                wsTemplate.Cells["A2"].PutValue("Not a number");
                wsTemplate.Cells["B1"].PutValue("678");
                wsTemplate.Cells["C5"].PutValue("Text");
                wbTemplate.Save(inputPath);
                Console.WriteLine($"Created placeholder workbook '{inputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create placeholder workbook: {ex.Message}");
                return;
            }
        }

        try
        {
            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a range that covers cells A1:C10 (use fully qualified type to avoid ambiguity)
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1", "C10");

            // Iterate through each cell in the range and convert numeric strings to numbers
            foreach (Cell cell in range)
            {
                if (cell.Type == CellValueType.IsString)
                {
                    if (double.TryParse(cell.StringValue, out double numericValue))
                    {
                        cell.PutValue(numericValue);
                    }
                }
            }

            // Count cells that are now numeric
            int convertedCount = 0;
            foreach (Cell cell in range)
            {
                if (cell.Type == CellValueType.IsNumeric)
                {
                    convertedCount++;
                }
            }

            Console.WriteLine($"Number of cells converted to numeric: {convertedCount}");

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Modified workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
