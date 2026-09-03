// Title: Retrieve the last populated row and freeze all rows above it using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an existing XLSX file with Aspose.Cells, determines the highest row containing data via the Cells.MaxDataRow property, applies Worksheet.FreezePanes to lock rows up to that index, and saves the modified workbook. | Demonstrate combining Worksheet.FreezePanes with Cells.MaxDataRow to lock the upper portion of a sheet in a .NET application. | Provide a .NET snippet that checks for the input file, calculates the maximum data row index, freezes rows from the first row through that index, and writes the result to a new XLSX file.
// Common Searches: Aspose.Cells C# get index of last row with data and freeze rows above | How to freeze panes up to the last populated row in an Excel workbook using Aspose.Cells for .NET | C# example for using MaxDataRow and FreezePanes to lock top rows in an XLSX file
// Tags: worksheet last row detection aspose.cells | freezepanes method c# | freeze top rows aspose.cells | excel row freezing .net | input xlsx output xlsx aspose.cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads input.xlsx, obtains the index of the final data row via worksheet.Cells.MaxDataRow, freezes all rows from the top through that index using FreezePanes, and saves the result as output.xlsx.
class Program
{
    static void Main()
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

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the zero‑based index of the last row that contains data
            int maxDataRowIndex = worksheet.Cells.MaxDataRow;

            // Freeze all rows above the maximum data row.
            // FreezePanes(row, column, totalRows, totalColumns) freezes rows up to totalRows-1 and columns up to totalColumns-1.
            // We freeze rows from 0 to maxDataRowIndex (inclusive) and no columns.
            worksheet.FreezePanes(maxDataRowIndex + 1, 0, maxDataRowIndex + 1, 0);

            // Save the modified workbook
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
