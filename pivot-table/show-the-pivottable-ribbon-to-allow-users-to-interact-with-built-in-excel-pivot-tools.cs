// Title: Show the Excel PivotTable Tools ribbon by enabling field list, dialog, and wizard in a C# Aspose.Cells workbook
// AI Prompts: Generate C# code with Aspose.Cells that creates a pivot table, assigns Region as a row field and Sales as a data field, then turns on the field list, field dialog, and wizard so the PivotTable Tools ribbon becomes visible in Excel. | Write a .NET example that adds a pivot table, refreshes and calculates its data, activates the UI features (EnableFieldList, EnableFieldDialog, EnableWizard), and saves the workbook with the PivotTable ribbon exposed.
// Common Searches: how to programmatically display the PivotTable Tools ribbon with Aspose.Cells for .NET | C# Aspose.Cells enable field list on a pivot table to show Excel UI | show Excel pivot table ribbon using EnableWizard property in Aspose.Cells | Aspose.Cells pivot table UI options EnableFieldDialog example | display pivot field list pane in generated Excel file using Aspose.Cells C#
// Tags: Aspose.Cells pivot UI activation C# | expose PivotTable ribbon via Aspose.Cells | configure pivot row and data fields Aspose.Cells | refresh calculate pivot cache Aspose.Cells | add pivot table to worksheet Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, fills it with Region and Sales data, adds a pivot table on A1:B5, places Region in the row area and Sales in the data area, enables the field list, field dialog, and wizard to make the PivotTable Tools ribbon appear, refreshes and calculates the pivot, and saves the file as ShowPivotTableRibbonDemo.xlsx.
    public class ShowPivotTableRibbonDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Region";
                sheet.Cells["B1"].Value = "Sales";
                sheet.Cells["A2"].Value = "North";
                sheet.Cells["B2"].Value = 1200;
                sheet.Cells["A3"].Value = "South";
                sheet.Cells["B3"].Value = 1500;
                sheet.Cells["A4"].Value = "East";
                sheet.Cells["B4"].Value = 800;
                sheet.Cells["A5"].Value = "West";
                sheet.Cells["B5"].Value = 950;

                // Add a pivot table to the worksheet
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure pivot fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Enable UI features that cause Excel to display the PivotTable ribbon
                pivotTable.EnableFieldList = true;   // Shows the field list pane and activates the PivotTable Tools ribbon
                pivotTable.EnableFieldDialog = true; // Allows the field dialog to be opened on double‑click
                pivotTable.EnableWizard = true;      // Makes the PivotTable Wizard available

                // Refresh and calculate the pivot data using the correct API
                pivotTable.RefreshData();   // Refreshes the pivot cache
                pivotTable.CalculateData(); // Calculates the pivot table values

                // Save the workbook
                workbook.Save("ShowPivotTableRibbonDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ShowPivotTableRibbonDemo.Run();
        }
    }
}
