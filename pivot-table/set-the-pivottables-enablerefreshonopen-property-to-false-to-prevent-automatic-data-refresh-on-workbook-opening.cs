// Title: Disable automatic PivotTable refresh on workbook open using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that creates a PivotTable and disables its RefreshDataOnOpeningFile property. | Demonstrate how to configure a PivotTable in Aspose.Cells so it does not auto‑refresh when the workbook is opened. | Provide a .NET example that prevents a PivotTable from refreshing on file load with Aspose.Cells.
// Common Searches: Aspose.Cells C# turn off pivot table auto refresh on opening | set RefreshDataOnOpeningFile false for PivotTable Aspose | prevent Excel pivot refresh on workbook load using Aspose.Cells | how to disable pivot auto refresh in .NET Excel file | example disabling pivot table refresh on open Aspose.Cells
// Tags: Aspose.Cells PivotTable RefreshDataOnOpeningFile | C# disable pivot auto refresh | Excel workbook pivot refresh control | Aspose.Cells set pivot refresh false | prevent pivot table auto refresh .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, adds sample product and sales data, inserts a PivotTable, disables its automatic refresh on opening by setting RefreshDataOnOpeningFile to false, and saves the file as PivotTableDisableRefreshOnOpenDemo.xlsx.
    public class PivotTableDisableRefreshOnOpenDemo
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(1000);
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(2000);
            worksheet.Cells["A4"].PutValue("Orange");
            worksheet.Cells["B4"].PutValue(3000);

            // Add a pivot table based on the data range
            int pivotIndex = worksheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Configure the pivot table (row and data fields)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Sales column

            // Disable automatic refresh when the workbook is opened
            pivotTable.RefreshDataOnOpeningFile = false;

            // Save the workbook to a file
            string outputPath = "PivotTableDisableRefreshOnOpenDemo.xlsx";
            workbook.Save(outputPath);
        }
    }
}
