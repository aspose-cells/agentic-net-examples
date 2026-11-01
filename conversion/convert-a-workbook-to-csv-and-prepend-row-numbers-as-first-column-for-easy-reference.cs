// Title: Add a leading row number column and export an Excel worksheet to CSV using Aspose.Cells for .NET
// Common Searches: how to insert a row index column before exporting Excel to CSV with Aspose.Cells C# | Aspose.Cells C# export first worksheet to CSV with row numbers | prepend sequential numbers to each row in CSV generated from .xlsx using Aspose.Cells | C# code to add leftmost column with row numbers and save as CSV via Aspose.Cells | convert .xlsx to .csv and include line numbers using Aspose.Cells library
// Tags: insert leftmost column for row index Aspose.Cells | populate sequential row numbers in worksheet cells | export first worksheet to CSV using Aspose.Cells | add leading index column before CSV export | save modified workbook as CSV in .NET

using System;
using Aspose.Cells;

// The example loads an .xlsx file, inserts a new leftmost column, fills it with incremental row numbers, and saves the first worksheet as a CSV file using Aspose.Cells for .NET.
class WorkbookToCsvWithRowNumbers
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (you can modify this to target a specific sheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a new column at the very left to hold row numbers
        sheet.Cells.InsertColumn(0);

        // Determine the last row that contains data
        int lastRow = sheet.Cells.MaxDataRow;

        // Populate the new first column with sequential row numbers (starting at 1)
        for (int row = 0; row <= lastRow; row++)
        {
            // PutValue writes the integer into the cell
            sheet.Cells[row, 0].PutValue(row + 1);
        }

        // Save the modified workbook as CSV (the first worksheet is exported by default)
        workbook.Save("output.csv", SaveFormat.Csv);
    }
}
