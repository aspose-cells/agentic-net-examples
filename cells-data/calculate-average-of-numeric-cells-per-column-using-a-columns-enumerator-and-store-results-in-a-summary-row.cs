// Title: Calculate column averages and add a labeled summary row using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to loop through each worksheet column, compute the average of numeric cells, and write the result into a new row labeled 'Average'. | Create a reusable function in Aspose.Cells that returns a dictionary of column indexes mapped to their average values and then inserts those averages into the bottom row of the sheet. | Modify the column‑average example to skip cells containing errors, apply bold formatting to the summary row, and save the workbook as an .xlsx file.
// Common Searches: aspocells calculate average for each column and add totals row | c# enumerate worksheet columns with Aspose.Cells to compute statistics | how to write a summary row with column averages using Aspose.Cells .NET | skip error cells when averaging columns in Excel with Aspose.Cells | add bold formatting to a summary row in Aspose.Cells workbook
// Tags: column average calculation Aspose.Cells | enumerate worksheet columns C# | write summary row Excel Aspose.Cells | numeric cell aggregation Aspose.Cells .NET | format summary row bold Aspose.Cells

using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample creates a 10x5 numeric grid, enumerates each column via the Columns enumerator, computes the average of numeric cells, writes the averages into a new row labeled "Average", applies optional formatting, and saves the workbook as ColumnAverages.xlsx.
class AveragePerColumn
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample numeric data (rows 0-9, columns 0-4)
        int dataRows = 10;
        int dataCols = 5;
        for (int row = 0; row < dataRows; row++)
        {
            for (int col = 0; col < dataCols; col++)
            {
                cells[row, col].PutValue(row * col + 1); // simple non‑zero numbers
            }
        }

        // Row index where the summary (average) will be written
        int summaryRow = dataRows;

        // Write a label for the summary row
        cells[summaryRow, 0].PutValue("Average");

        // Enumerate through all instantiated columns
        IEnumerator colEnum = cells.Columns.GetEnumerator();
        while (colEnum.MoveNext())
        {
            // Each item is a Column object
            Column column = (Column)colEnum.Current;
            int colIndex = column.Index; // column index in the worksheet

            double sum = 0;
            int count = 0;

            // Iterate through the data rows of the current column
            for (int row = 0; row < dataRows; row++)
            {
                Cell cell = cells[row, colIndex];
                if (cell != null && cell.Type == CellValueType.IsNumeric)
                {
                    sum += cell.DoubleValue;
                    count++;
                }
            }

            // Calculate average and store it in the summary row
            double average = count > 0 ? sum / count : 0;
            cells[summaryRow, colIndex].PutValue(average);
        }

        // Save the workbook
        workbook.Save("ColumnAverages.xlsx");
    }
}
