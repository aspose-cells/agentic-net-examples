using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source workbook (create & load step)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Create a DataTable to map worksheet cells
        DataTable dataTable = new DataTable();

        // Determine the used range
        int maxRow = cells.MaxRow + 1;      // rows are zero‑based
        int maxColumn = cells.MaxColumn + 1;

        // Assume the first row contains column headers
        for (int col = 0; col < maxColumn; col++)
        {
            string header = cells[0, col].StringValue;
            if (string.IsNullOrEmpty(header))
                header = $"Column{col + 1}";
            dataTable.Columns.Add(header, typeof(string));
        }

        // Populate the DataTable with the remaining rows
        for (int row = 1; row < maxRow; row++)
        {
            DataRow dr = dataTable.NewRow();
            bool hasData = false;
            for (int col = 0; col < maxColumn; col++)
            {
                string value = cells[row, col].StringValue;
                dr[col] = value;
                if (!string.IsNullOrEmpty(value))
                    hasData = true;
            }
            // Add only rows that contain at least one value
            if (hasData)
                dataTable.Rows.Add(dr);
        }

        // Example modification: add a new row to the DataTable
        DataRow newRow = dataTable.NewRow();
        for (int i = 0; i < dataTable.Columns.Count; i++)
        {
            newRow[i] = $"NewValue{i + 1}";
        }
        dataTable.Rows.Add(newRow);

        // Clear existing worksheet data (optional, prepares for re‑import)
        worksheet.Cells.Clear();

        // Write headers back to the worksheet
        for (int col = 0; col < dataTable.Columns.Count; col++)
        {
            cells[0, col].PutValue(dataTable.Columns[col].ColumnName);
        }

        // Import each DataRow into the worksheet starting at row 1 (index 1)
        int targetRow = 1;
        foreach (DataRow dr in dataTable.Rows)
        {
            cells.ImportDataRow(dr, targetRow, 0);
            targetRow++;
        }

        // Save the modified workbook (save step)
        workbook.Save("output.xlsx");
    }
}