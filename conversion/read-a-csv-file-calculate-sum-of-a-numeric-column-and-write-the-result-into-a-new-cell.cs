// Title: Calculate column sum from a CSV file and write the total to a new cell with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells read CSV and sum values in a specific column | How to add a total row after CSV data using Aspose.Cells | Aspose.Cells .NET calculate numeric column total from CSV | Write computed sum to a new cell below data with Aspose.Cells C# | Convert CSV to Excel and include column sum using Aspose.Cells
// Tags: calculate column sum Aspose.Cells C# | write total to new cell Aspose.Cells | load CSV into workbook Aspose.Cells | save workbook as Excel Aspose.Cells | iterate rows numeric cells Aspose.Cells

using System;
using Aspose.Cells;

// The example loads a CSV file into an Aspose.Cells Workbook, iterates through a chosen column to sum numeric values, writes the sum and a label in a new row beneath the data, and saves the result as an Excel workbook.
class CsvSumExample
{
    static void Main()
    {
        // Load the CSV file into a workbook
        Workbook workbook = new Workbook("input.csv");

        // Get the first worksheet (CSV data is loaded into the first sheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Define the zero‑based index of the column to sum (e.g., column B => index 1)
        int columnToSum = 1;

        // Calculate the sum of all numeric cells in the specified column
        double sum = 0.0;
        // MaxDataRow returns the last row that contains data
        for (int row = 0; row <= sheet.Cells.MaxDataRow; row++)
        {
            Cell cell = sheet.Cells[row, columnToSum];
            if (cell.Type == CellValueType.IsNumeric)
            {
                sum += cell.DoubleValue;
            }
        }

        // Write the sum into a new cell below the existing data
        int resultRow = sheet.Cells.MaxDataRow + 2; // leave one empty row
        sheet.Cells[resultRow, columnToSum].PutValue(sum);

        // Optionally, add a label for the result in the adjacent column
        sheet.Cells[resultRow, columnToSum - 1].PutValue("Sum:");

        // Save the workbook to an Excel file (you can also save as CSV if needed)
        workbook.Save("output.xlsx");
    }
}
