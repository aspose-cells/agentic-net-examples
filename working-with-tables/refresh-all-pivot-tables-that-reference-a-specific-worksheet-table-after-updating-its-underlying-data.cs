// Title: Refresh All Pivot Tables After Updating a Worksheet Table with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to load or create an Excel workbook, locate a worksheet and a ListObject (table), modify a cell in the table, and call Worksheets.RefreshPivotTables() to update every pivot table that uses the table before saving the file. Includes fallback handling for missing files, worksheets, or tables.
// Keywords: Aspose.Cells | RefreshPivotTables | C# | .NET | ListObject | Excel table update | pivot cache refresh | programmatic pivot table refresh | worksheet table modification | Excel automation
// Common Searches: Aspose.Cells refresh all pivot tables after table change | C# update ListObject and refresh pivot tables | Worksheets.RefreshPivotTables example | how to refresh pivot tables linked to a table in Aspose.Cells | programmatically refresh pivot cache .NET
// Developer Intent: Programmatically refresh every pivot table that depends on a modified worksheet table.
// Use Cases: Update data in a ListObject and ensure all related pivot reports are current before exporting the workbook. | Create a workbook with a table, change its values via code, and automatically synchronize all pivot tables. | Load an existing Excel file, edit table rows, and call RefreshPivotTables to keep the pivot cache consistent.
// AI Prompts: Generate C# code that changes a cell in a ListObject and then refreshes all dependent pivot tables using Aspose.Cells. | Show how to safely locate a worksheet and a ListObject by name, modify its data, and invoke Worksheets.RefreshPivotTables in .NET. | Provide an example that handles missing workbook, worksheet, or table while still performing a pivot table refresh.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables; // for ListObject

// Demonstrates how to load or create an Excel workbook, locate a worksheet and a ListObject (table), modify a cell in the table, and call Worksheets.RefreshPivotTables() to update every pivot table that uses the table before saving the file. Includes fallback handling for missing files, worksheets, or tables.
class RefreshPivotTablesDemo
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; create a simple workbook if missing
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                var ws = tempWb.Worksheets[0];
                ws.Name = "Data";

                // Sample data
                ws.Cells["A1"].PutValue("Header1");
                ws.Cells["B1"].PutValue("Header2");
                ws.Cells["A2"].PutValue("Value1");
                ws.Cells["B2"].PutValue("Value2");

                // Create a table named MyTable covering the sample data
                var listObj = ws.ListObjects[ws.ListObjects.Add(0, 0, 1, 1, true)];
                listObj.DisplayName = "MyTable";

                tempWb.Save(inputPath);
            }

            // Load the workbook that contains the data table and pivot tables
            var workbook = new Workbook(inputPath);

            // Access the worksheet that holds the source table (fallback to first sheet if not found)
            Worksheet dataSheet = workbook.Worksheets["Data"] ?? workbook.Worksheets[0];

            // Access the table by its name (fallback to first table if not found)
            ListObject table = dataSheet.ListObjects["MyTable"];
            if (table == null && dataSheet.ListObjects.Count > 0)
                table = dataSheet.ListObjects[0];

            if (table != null)
            {
                // Modify the first data row of the first column in the table
                int firstDataRow = table.DataRange.FirstRow;
                int firstDataColumn = table.DataRange.FirstColumn;
                dataSheet.Cells[firstDataRow, firstDataColumn].PutValue("UpdatedValue");
            }
            else
            {
                Console.WriteLine("No ListObject found in the worksheet.");
            }

            // Refresh all pivot tables in the workbook so they reflect the updated table data
            workbook.Worksheets.RefreshPivotTables();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
