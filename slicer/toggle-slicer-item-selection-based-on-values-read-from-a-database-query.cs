using System;
using System.Data;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];
            Cells cells = ws.Cells;

            // Retrieve data from a database (simulated here)
            DataTable dataTable = GetDataFromDatabase();

            // Populate the worksheet with the retrieved data starting at cell A1
            for (int c = 0; c < dataTable.Columns.Count; c++)
            {
                cells[0, c].PutValue(dataTable.Columns[c].ColumnName);
            }
            for (int r = 0; r < dataTable.Rows.Count; r++)
            {
                for (int c = 0; c < dataTable.Columns.Count; c++)
                {
                    cells[r + 1, c].PutValue(dataTable.Rows[r][c]);
                }
            }

            // Define the data range for the pivot table (e.g., "A1:B5")
            string lastColumnLetter = GetColumnLetter(dataTable.Columns.Count - 1);
            string dataRange = $"A1:{lastColumnLetter}{dataTable.Rows.Count + 1}";

            // Add a pivot table based on the data range
            int pivotIdx = ws.PivotTables.Add(dataRange, "E5", "PivotTable1");
            PivotTable pivot = ws.PivotTables[pivotIdx];
            // Use the first column as the row field (the slicer will be based on this field)
            pivot.AddFieldToArea(PivotFieldType.Row, 0);
            // If there is a second column, add it as a data field
            if (dataTable.Columns.Count > 1)
                pivot.AddFieldToArea(PivotFieldType.Data, 1);
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer for the first field of the pivot table
            int slicerIdx = ws.Slicers.Add(pivot, "G5", dataTable.Columns[0].ColumnName);
            Slicer slicer = ws.Slicers[slicerIdx];

            // Retrieve the list of values that should be selected in the slicer (simulated)
            HashSet<string> valuesToSelect = GetSelectedValuesFromDatabase();

            // Toggle slicer items: select those whose Value exists in valuesToSelect, deselect others
            foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
            {
                item.Selected = valuesToSelect.Contains(item.Value);
            }

            // Refresh the slicer to apply the selection changes
            slicer.Refresh();

            // Save the workbook
            string outputPath = "SlicerSelectionDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper: converts a zero‑based column index to an Excel column letter (0 -> "A")
    static string GetColumnLetter(int columnIndex)
    {
        int dividend = columnIndex + 1;
        string columnName = string.Empty;
        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }
        return columnName;
    }

    // Simulated database query that returns the data for the worksheet
    static DataTable GetDataFromDatabase()
    {
        // Replace this stub with actual database access code (e.g., SqlConnection, SqlCommand)
        DataTable table = new DataTable();
        table.Columns.Add("Category", typeof(string));
        table.Columns.Add("Amount", typeof(int));
        table.Rows.Add("Apple", 120);
        table.Rows.Add("Banana", 80);
        table.Rows.Add("Orange", 150);
        table.Rows.Add("Apple", 90);
        return table;
    }

    // Simulated database query that returns the slicer values that must be selected
    static HashSet<string> GetSelectedValuesFromDatabase()
    {
        // Replace this stub with actual database access code
        return new HashSet<string> { "Apple", "Orange" };
    }
}