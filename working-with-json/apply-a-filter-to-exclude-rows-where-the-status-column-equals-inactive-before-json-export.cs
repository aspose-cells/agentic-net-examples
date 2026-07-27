using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data with a "Status" column
        worksheet.Cells["A1"].PutValue("ID");
        worksheet.Cells["B1"].PutValue("Name");
        worksheet.Cells["C1"].PutValue("Status");

        worksheet.Cells["A2"].PutValue(1);
        worksheet.Cells["B2"].PutValue("Alice");
        worksheet.Cells["C2"].PutValue("Active");

        worksheet.Cells["A3"].PutValue(2);
        worksheet.Cells["B3"].PutValue("Bob");
        worksheet.Cells["C3"].PutValue("Inactive");

        worksheet.Cells["A4"].PutValue(3);
        worksheet.Cells["B4"].PutValue("Charlie");
        worksheet.Cells["C4"].PutValue("Active");

        // Set the auto‑filter range (including header row)
        // Parameters: startRow, startColumn, endRow
        worksheet.AutoFilter.SetRange(0, 0, 3);

        // Apply a filter on the "Status" column (field index 2) to keep only "Active" rows
        worksheet.AutoFilter.AddFilter(2, "Active");
        worksheet.AutoFilter.Refresh(); // Hide rows that do not match the filter (i.e., "Inactive")

        // Configure JSON save options (skip empty rows for cleaner output)
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            SkipEmptyRows = true
        };

        // Save the workbook to JSON; hidden rows (Inactive) will not be exported
        workbook.Save("FilteredData.json", jsonOptions);
    }
}