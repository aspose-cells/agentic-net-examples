using System;
using Aspose.Cells;

class WorkbookToCsvWithRowNumbers
{
    static void Main()
    {
        // Path to the source Excel workbook
        string sourcePath = "input.xlsx";

        // Path for the resulting CSV file
        string csvPath = "output.csv";

        // Load the workbook
        Workbook workbook = new Workbook(sourcePath);

        // Access the first worksheet (you can modify this to target a specific sheet)
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Insert a new column at the very left (index 0) to hold row numbers
        cells.InsertColumn(0);

        // Determine the last row that contains data
        int lastRow = cells.MaxDataRow; // zero‑based index

        // Populate the new column with sequential row numbers (starting at 1)
        for (int row = 0; row <= lastRow; row++)
        {
            cells[row, 0].PutValue(row + 1);
        }

        // Save the modified workbook as CSV
        workbook.Save(csvPath, SaveFormat.Csv);
    }
}