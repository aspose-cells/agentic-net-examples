// Title: Insert a total row with SUM formulas for columns M and N using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add a total row after the last data row | C# code to sum columns M and N in an Excel file with Aspose.Cells | How to generate SUM formulas for multiple columns using Aspose.Cells .NET | Insert subtotal row for specific columns in Excel via Aspose.Cells C# | Programmatically calculate column totals in .xlsx with Aspose.Cells library
// Tags: Aspose.Cells insert total row C# | SUM formula generation Aspose.Cells | calculate column M N totals Aspose.Cells | add subtotal for multiple columns .NET | Excel worksheet last data row detection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, finds the last populated row, inserts a new row below it, writes a "Total" label, and sets SUM formulas for columns M and N (indices 12 and 13) that sum the data from row 2 to the last data row, then saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Columns to be summed: M (index 12) and N (index 13)
            int[] totalColumns = new int[] { 12, 13 };

            // Determine the last row that contains data
            int lastDataRow = worksheet.Cells.MaxDataRow;
            if (lastDataRow < 0)
            {
                Console.WriteLine("The worksheet does not contain any data.");
                return;
            }

            // Insert a new row after the last data row for totals
            int totalRowIndex = lastDataRow + 1;
            worksheet.Cells.InsertRows(totalRowIndex, 1);

            // Add a label for the totals row (optional)
            worksheet.Cells[totalRowIndex, 0].PutValue("Total");

            // Set SUM formulas for each specified column
            foreach (int colIndex in totalColumns)
            {
                string columnName = CellsHelper.ColumnIndexToName(colIndex);
                // Assuming data starts at row 2 (index 1) after a header row
                string formula = $"SUM({columnName}2:{columnName}{lastDataRow + 1})";
                worksheet.Cells[totalRowIndex, colIndex].Formula = formula;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
