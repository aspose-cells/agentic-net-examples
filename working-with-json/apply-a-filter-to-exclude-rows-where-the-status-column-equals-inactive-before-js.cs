using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Utility;

class ExportActiveRowsToJson
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data (Header + rows)
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("Name");
        cells["C1"].PutValue("Status");

        cells["A2"].PutValue(1);
        cells["B2"].PutValue("Alice");
        cells["C2"].PutValue("Active");

        cells["A3"].PutValue(2);
        cells["B3"].PutValue("Bob");
        cells["C3"].PutValue("Inactive");

        cells["A4"].PutValue(3);
        cells["B4"].PutValue("Charlie");
        cells["C4"].PutValue("Active");

        // Define the auto‑filter range (including header and data rows)
        sheet.AutoFilter.SetRange(0, 0, 3);

        // Apply filter on the Status column (index 2) to keep only "Active" rows
        sheet.AutoFilter.AddFilter(2, "Active");
        sheet.AutoFilter.Refresh();

        // Export only the visible rows to a DataTable
        ExportTableOptions exportOptions = new ExportTableOptions
        {
            PlotVisibleRows = true,
            ExportColumnName = true
        };

        DataTable dt = cells.ExportDataTable(0, 0, 4, 3, exportOptions);

        // Convert DataTable to a serializable list of dictionaries
        var rows = new List<Dictionary<string, object>>();
        foreach (DataRow dr in dt.Rows)
        {
            var dict = new Dictionary<string, object>();
            foreach (DataColumn col in dt.Columns)
            {
                dict[col.ColumnName] = dr[col];
            }
            rows.Add(dict);
        }

        // Serialize to JSON
        string json = JsonSerializer.Serialize(rows, new JsonSerializerOptions { WriteIndented = true });

        // Save the JSON string to a file
        string jsonPath = "ActiveRows.json";
        File.WriteAllText(jsonPath, json);

        Console.WriteLine($"JSON exported to '{jsonPath}':");
        Console.WriteLine(json);
    }
}