// Title: Batch add slicers with a uniform column width to multiple Excel workbooks using Aspose.Cells for .NET (C#)
// Description: A C# example that iterates over a list of workbook files, loads each workbook, adds a slicer for the first pivot‑table field (or creates a ListObject and adds a slicer when no pivot exists), sets the slicer's ColumnWidth to a common value, and saves the workbook. Demonstrates batch processing and consistent slicer sizing across many Excel files.
// Keywords: Aspose.Cells | C# slicer | batch add slicer | uniform slicer width | Excel pivot slicer | ListObject slicer | process multiple workbooks | Aspose.Cells .NET | slicer ColumnWidth | automate Excel formatting
// Common Searches: How to add a slicer to several Excel files with Aspose.Cells | Set the same slicer width for multiple workbooks in C# | Create slicer from a pivot table using Aspose.Cells .NET | Add slicer to a ListObject when no pivot table exists | Batch update slicer properties with Aspose.Cells
// Developer Intent: Add a slicer to each workbook and enforce a consistent column width across all slicers.
// Use Cases: Standardize slicer appearance in a suite of financial dashboards before distribution. | Generate slicers for tables in workbooks that lack pivot tables while keeping layout uniform. | Automate the preparation of Excel reports for a multinational team, ensuring slicer size matches corporate UI guidelines.
// AI Prompts: Generate a reusable C# method that accepts a collection of workbook paths and a column width, then adds slicers to each workbook handling both pivot tables and ListObjects with Aspose.Cells. | Provide error‑handling code for missing files, empty worksheets, or invalid data ranges when batch adding slicers. | Show how to also set slicer style, position, and column width uniformly across multiple workbooks using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

// A C# example that iterates over a list of workbook files, loads each workbook, adds a slicer for the first pivot‑table field (or creates a ListObject and adds a slicer when no pivot exists), sets the slicer's ColumnWidth to a common value, and saves the workbook. Demonstrates batch processing and consistent slicer sizing across many Excel files.
class BatchSlicerProcessor
{
    static void Main()
    {
        // Paths of workbooks to process
        string[] workbookFiles = new string[]
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx",
            "Workbook3.xlsx"
        };

        // Desired uniform column width for all slicers (points)
        double uniformColumnWidth = 80.0;

        foreach (string filePath in workbookFiles)
        {
            // Load existing workbook
            Workbook workbook = new Workbook(filePath);

            // Work with the first worksheet (adjust as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Try to add a slicer based on an existing pivot table
            if (sheet.PivotTables.Count > 0)
            {
                // Use the first pivot table in the sheet
                PivotTable pivot = sheet.PivotTables[0];

                // Use the first base field of the pivot table for the slicer
                string baseFieldName = pivot.BaseFields[0].Name;

                // Add slicer at cell A1 (upper‑left corner of slicer range)
                int slicerIndex = sheet.Slicers.Add(pivot, "A1", baseFieldName);
                Slicer slicer = sheet.Slicers[slicerIndex];

                // Apply uniform column width
                slicer.ColumnWidth = uniformColumnWidth;
            }
            else
            {
                // No pivot table – create a simple table (ListObject) from a data range
                // Assumes data exists in A1:B5; adjust range as required
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Add slicer for the first column of the table at cell D1
                int slicerIndex = sheet.Slicers.Add(table, 0, "D1");
                Slicer slicer = sheet.Slicers[slicerIndex];

                // Apply uniform column width
                slicer.ColumnWidth = uniformColumnWidth;
            }

            // Save the workbook (overwrites the original file)
            workbook.Save(filePath);
        }
    }
}
