// Title: Create an Excel slicer linked to a table column and a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that builds a workbook, defines a ListObject, adds a column chart based on the table data, and inserts a slicer that filters the chart by the Category field. | Generate a .NET program that populates a worksheet with sample data, creates an Excel table, attaches a column chart to the table range, and adds a slicer connected to the table's Category column.
// Common Searches: Aspose.Cells C# how to add a slicer that controls a chart linked to a table column | example code for creating an Excel slicer for dynamic chart filtering with Aspose.Cells .NET | C# Aspose.Cells add ListObject and slicer to filter column chart | programmatically insert a slicer for a table column and connect it to a chart using Aspose.Cells
// Tags: Aspose.Cells add slicer to worksheet | C# create Excel table ListObject | Aspose.Cells column chart from ListObject | slicer linked to table column Aspose.Cells | dynamic chart filtering Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Charts;
using Aspose.Cells.Slicers;   // Required for Slicer support

// The example creates a new workbook, adds a ListObject with Category and Value data, generates a column chart that references the table range, inserts a slicer linked to the Category column, and saves the file as SlicerChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and name it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate data for the table
            // Header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            // Sample data rows
            string[] categories = { "A", "B", "C", "D" };
            int[] values = { 10, 20, 30, 40 };
            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(categories[i]); // Column A
                sheet.Cells[i + 1, 1].PutValue(values[i]);    // Column B
            }

            // Define the range that will become a table (including header)
            int firstRow = 0;
            int firstColumn = 0;
            int totalRows = categories.Length + 1; // + header row
            int totalColumns = 2;

            // Add a ListObject (Excel table) to the worksheet
            int tableIdx = sheet.ListObjects.Add(firstRow, firstColumn,
                firstRow + totalRows - 1, firstColumn + totalColumns - 1, true);
            ListObject table = sheet.ListObjects[tableIdx];

            // Optional: set a display name if supported
            if (table.GetType().GetProperty("DisplayName") != null)
                table.DisplayName = "DataTable";

            table.ShowHeaderRow = true;
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIdx];
            chart.Title.Text = "Category Values";

            // Add a series that uses the table range
            int seriesIdx = chart.NSeries.Add(table.DataRange.RefersTo, true);

            // Build address strings for X (Category) and Y (Value) ranges (excluding header)
            string sheetName = sheet.Name;
            string catColumnLetter = CellsHelper.ColumnIndexToName(table.DataRange.FirstColumn);
            string valColumnLetter = CellsHelper.ColumnIndexToName(table.DataRange.FirstColumn + 1);
            int dataStartRow = table.DataRange.FirstRow + 2; // +2 to skip header (1‑based Excel rows)
            int dataEndRow = table.DataRange.FirstRow + table.DataRange.RowCount;

            chart.NSeries[seriesIdx].XValues = $"{sheetName}!{catColumnLetter}{dataStartRow}:{catColumnLetter}{dataEndRow}";
            chart.NSeries[seriesIdx].Values = $"{sheetName}!{valColumnLetter}{dataStartRow}:{valColumnLetter}{dataEndRow}";
            chart.NSeries[seriesIdx].Name = "Values";

            // Add a slicer linked to the "Category" column of the table
            int slicerIdx = sheet.Slicers.Add(table, 0, "Category");
            Slicer slicer = sheet.Slicers[slicerIdx];
            slicer.Name = "CategorySlicer";
            slicer.Caption = "Category";

            // Save the workbook
            string outputPath = "SlicerChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
