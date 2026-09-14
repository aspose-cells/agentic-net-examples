// Title: Disable Excel 2003 compatibility and refresh every PivotTable in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads a workbook, disables Excel 2003 compatibility for each PivotTable, refreshes them, and saves the result using Aspose.Cells. | Demonstrate iterating through all worksheets to turn off PivotTable.IsExcel2003Compatible and then invoke the appropriate refresh method in Aspose.Cells.
// Common Searches: Aspose.Cells C# how to turn off Excel2003 compatibility for pivot tables and refresh data | refresh all pivot tables in a workbook after disabling Excel 2003 mode using Aspose.Cells | programmatic way to update pivot tables when compatibility flag is changed in .NET | example code to set PivotTable.IsExcel2003Compatible false and refresh workbook with Aspose.Cells | C# Aspose.Cells refresh pivot tables across all worksheets
// Tags: PivotTable.IsExcel2003Compatible property Aspose.Cells | RefreshPivotTables API C# | iterate worksheets refresh pivot tables Aspose.Cells | disable Excel2003 compatibility for pivot tables .NET | update pivot tables after compatibility change Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshExample
{
    // Loads an existing XLSX file, disables Excel 2003 compatibility for every PivotTable on each worksheet, refreshes all pivot tables, and saves the updated workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains one or more PivotTables
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Check if the worksheet contains any PivotTables
                if (sheet.PivotTables.Count > 0)
                {
                    // Disable Excel 2003 compatibility for each PivotTable
                    foreach (PivotTable pt in sheet.PivotTables)
                    {
                        pt.IsExcel2003Compatible = false;
                    }

                    // Refresh all PivotTables in the current worksheet
                    sheet.RefreshPivotTables();
                }
            }

            // Alternatively, refresh all PivotTables in the entire workbook at once
            // workbook.Worksheets.RefreshPivotTables();

            // Save the updated workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
