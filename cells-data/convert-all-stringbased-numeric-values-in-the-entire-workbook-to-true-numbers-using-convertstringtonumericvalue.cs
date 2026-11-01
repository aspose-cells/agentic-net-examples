// Title: Convert string‑based numeric values to true numbers in every cell of an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that walks through all worksheets and changes cells containing numeric strings into actual numeric types. | Enhance the conversion logic to handle numeric strings with commas or other thousand separators before saving the workbook. | Add a reporting feature that counts and prints the number of cells converted from string to numeric during processing.
// Common Searches: asp.net convert numeric strings to numbers in all Excel cells using Aspose.Cells | c# iterate over entire workbook replace string numbers with doubles | bulk conversion of string cells to numeric values Aspose.Cells | change cell type from string to numeric Aspose.Cells .NET example | process Excel workbook to transform text numbers into real numbers with Aspose.Cells
// Tags: convert string cells to numeric Aspose.Cells | bulk numeric conversion Excel workbook C# | parse numeric strings with thousand separators Aspose.Cells | replace text numbers with double values in Excel | save workbook after type conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, iterates through each worksheet and every cell, checks for cells of type string, attempts to parse the string as a double, and if successful replaces the cell content with the numeric value before saving the modified workbook to the specified output path.
class Program
{
    static void Main()
    {
        string inputPath = "{InputPath}";
        string outputPath = "{OutputPath}";

        // Ensure the input file exists before loading
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all cells in the worksheet
                foreach (Cell cell in sheet.Cells)
                {
                    // Process only string cells
                    if (cell.Type == CellValueType.IsString)
                    {
                        // Try to parse the string as a double
                        if (double.TryParse(cell.StringValue, out double numericValue))
                        {
                            // Replace the string with a numeric value
                            cell.PutValue(numericValue);
                        }
                    }
                }
            }

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
