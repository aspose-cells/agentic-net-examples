// Title: Concatenate raw string values from a specific Excel range and write the result to a summary cell using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook, iterate over cells A1:C3, collect only string values, concatenate them, and store the combined text in cell D1 with Aspose.Cells in C#. | Create a C# script that uses Aspose.Cells to define a range, filter cells by CellValueType.IsString, build a single string from those values, and write it to a designated summary cell. | Write C# code that reads raw strings from a given range, merges them into one string, places the result in another cell, and saves the workbook using Aspose.Cells.
// Common Searches: how to concatenate text from a range of cells using Aspose.Cells in C# | Aspose.Cells read only string cells from A1:C3 and combine them into D1 | C# example for merging Excel cell strings and saving the workbook with Aspose.Cells
// Tags: concatenate string values range Aspose.Cells C# | filter cells by CellValueType.IsString Aspose.Cells | output combined text to cell D1 Aspose.Cells | create range and iterate cells Aspose.Cells | save updated Excel workbook Aspose.Cells .NET

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Loads an Excel workbook (or creates a new one), reads raw string values from the range A1:C3 on the first worksheet, concatenates them, writes the combined text to cell D1, and saves the workbook to the specified output file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Paths to the input and output Excel files
            string inputPath = @"C:\Data\Input.xlsx";
            string outputPath = @"C:\Data\Output.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Load the workbook or create a new one if the input file is missing
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range to read raw strings from (e.g., A1:C3)
            string rangeAddress = "A1:C3";

            // Explicitly use Aspose.Cells.Range to avoid conflict with System.Range
            Aspose.Cells.Range range = sheet.Cells.CreateRange(rangeAddress);

            // Concatenate raw string values from the range
            StringBuilder concatenated = new StringBuilder();

            foreach (Cell cell in range)
            {
                if (cell.Type == CellValueType.IsString)
                {
                    concatenated.Append(cell.StringValue);
                }
            }

            // Write the concatenated result to a summary cell (e.g., D1)
            sheet.Cells["D1"].PutValue(concatenated.ToString());

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
