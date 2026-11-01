// Title: Compute average of numeric cells per column and write results to a summary row using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates over each column, calculates the average of numeric cells, and places the average in a new row below the data. | Adjust the snippet to skip non‑numeric cells and empty rows while computing column averages in an Excel workbook with Aspose.Cells. | Create a reusable C# method that returns a dictionary of column indexes to their average numeric values and optionally writes those averages to a designated summary row using Aspose.Cells.
// Common Searches: Aspose.Cells C# calculate column averages and add a summary row at the bottom of the sheet | How to ignore non‑numeric cells when computing per‑column averages with Aspose.Cells .NET | C# example enumerating columns in Aspose.Cells and writing summary statistics | Save Excel file with a new row containing per‑column averages using Aspose.Cells
// Tags: column average calculation Aspose.Cells .NET | write summary row Excel Aspose.Cells | enumerate columns Aspose.Cells API | numeric cell aggregation Aspose.Cells | per‑column statistics Aspose.Cells | C# Excel average per column Aspose.Cells

using Aspose.Cells;
using System;

// Load an existing workbook (replace with your actual file path)
Workbook workbook = new Workbook("input.xlsx");

// Get the first worksheet
Worksheet sheet = workbook.Worksheets[0];
Cells cells = sheet.Cells;

// Determine the row index for the summary (one row after the last used row)
int summaryRowIndex = cells.MaxDataRow + 1;

// Enumerate through all columns that contain data
foreach (Column column in cells.Columns)
{
    double sum = 0;
    int count = 0;

    // Iterate over each cell in the current column up to the last used row
    for (int row = 0; row <= cells.MaxDataRow; row++)
    {
        Cell cell = cells[row, column.Index];
        if (cell.Type == CellValueType.IsNumeric)
        {
            sum += cell.DoubleValue;
            count++;
        }
    }

    // Calculate average if there are numeric values
    double average = count > 0 ? sum / count : 0;

    // Store the average in the summary row, same column
    cells[summaryRowIndex, column.Index].PutValue(average);
}

// Save the workbook with the summary row added (replace with your desired output path)
workbook.Save("output.xlsx");
