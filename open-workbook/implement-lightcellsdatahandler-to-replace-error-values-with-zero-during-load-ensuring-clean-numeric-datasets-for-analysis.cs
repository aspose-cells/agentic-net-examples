// Title: Replace Excel error cells with zero using Aspose.Cells for .NET (C#)
// Common Searches: how to convert #DIV/0! and #N/A errors to 0 in an Excel file with Aspose.Cells C# | Aspose.Cells C# iterate through workbook cells and set error values to zero | clean numeric dataset by removing error cells during Excel import using Aspose.Cells | C# code example for replacing error cell values with numeric zero in Aspose.Cells | sanitize Excel data by handling error cells in Aspose.Cells .NET
// Tags: replace error cells with zero Aspose.Cells | clean numeric data Excel Aspose.Cells C# | iterate workbook cells error handling Aspose.Cells | sanitize Excel error values during load .NET | Aspose.Cells data cleaning example

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, scans every worksheet and cell, replaces any cell whose type is an error with the numeric value 0, and saves the cleaned workbook to a new file.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output_cleaned.xlsx";

        try
        {
            // Verify that the input file exists.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook.
            var workbook = new Workbook(inputPath);

            // Replace error values with numeric zero.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (cell.Type == CellValueType.IsError)
                    {
                        cell.PutValue(0.0);
                    }
                }
            }

            // Save the cleaned workbook.
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
