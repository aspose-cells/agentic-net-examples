// Title: Sum a numeric column in a CSV file and write the total to a new cell with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells calculate total of a column in a CSV workbook | How to add a sum row after CSV data using Aspose.Cells .NET | Write aggregated value to a new cell in Excel after loading CSV with Aspose.Cells | Aspose.Cells iterate over CSV rows to compute column sum in C# | Export CSV to XLSX with summed column value using Aspose.Cells
// Tags: column sum calculation Aspose.Cells C# | insert result cell after data Aspose.Cells | CSV to XLSX conversion with total row Aspose.Cells | load CSV workbook LoadOptions Aspose.Cells | aggregate numeric column Aspose.Cells worksheet

using Aspose.Cells;
using System;

// The example loads a CSV file into an Aspose.Cells Workbook using LoadOptions, iterates through a specified numeric column to compute the sum while skipping non‑numeric cells, writes a label and the calculated total a couple of rows below the existing data, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        // Path to the source CSV file
        string csvPath = "input.csv";

        // Load the CSV file into a workbook (Aspose.Cells automatically parses CSV)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
        Workbook workbook = new Workbook(csvPath, loadOptions);

        // Get the first worksheet (the CSV data is placed here)
        Worksheet sheet = workbook.Worksheets[0];

        // Index of the numeric column to sum (0‑based). Example: column B => index 1
        int numericColumnIndex = 1;

        // Calculate the sum of the numeric column
        double sum = 0;
        // Determine the last row that contains data
        int lastDataRow = sheet.Cells.MaxDataRow;
        for (int row = 0; row <= lastDataRow; row++)
        {
            // Try to parse the cell value as a double; ignore non‑numeric cells
            double value;
            if (double.TryParse(sheet.Cells[row, numericColumnIndex].StringValue, out value))
            {
                sum += value;
            }
        }

        // Write the result into a new cell.
        // Place a label in column A and the sum in column B, two rows after the data.
        int resultRow = lastDataRow + 2;
        sheet.Cells[resultRow, 0].PutValue("Sum");
        sheet.Cells[resultRow, 1].PutValue(sum);

        // Save the workbook to an XLSX file (or any format you need)
        workbook.Save("output.xlsx");
    }
}
