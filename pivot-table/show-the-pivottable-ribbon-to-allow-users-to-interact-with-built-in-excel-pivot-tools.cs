// Title: Enable PivotTable Ribbon (Field List, Dialog, Wizard) in Excel using Aspose.Cells for .NET
// Description: Creates a workbook, inserts sample sales data, builds a PivotTable on a separate sheet, assigns row and data fields, and turns on the PivotTable UI options—field list pane, field dialog, and wizard—before refreshing, calculating, and saving the file as PivotTable_With_Ribbon.xlsx.
// Keywords: Aspose.Cells | .NET | PivotTable ribbon | field list | field dialog | wizard | programmatic Excel pivot | enable PivotTable UI | Aspose.Cells example
// Common Searches: how to show pivot field list in Aspose.Cells .NET | enable pivot wizard programmatically with Aspose.Cells | display pivot table ribbon tools in generated Excel file | Aspose.Cells enable field dialog for PivotTable | C# code to activate PivotTable UI features
// Developer Intent: Expose the Excel PivotTable UI (field list, dialog, wizard) in a workbook generated with Aspose.Cells.
// Use Cases: Deliver a sales dashboard where analysts can rearrange rows, columns, and values using the native ribbon controls. | Provide an Excel template with a pre‑configured pivot and full UI access for end‑users to explore data without additional coding. | Integrate an interactive pivot table into a web‑based reporting portal by enabling the built‑in ribbon features in the exported file.
// AI Prompts: Generate C# code with Aspose.Cells that creates a PivotTable and turns on the field list, field dialog, and wizard UI. | Explain the purpose of EnableFieldList, EnableFieldDialog, and EnableWizard properties in Aspose.Cells PivotTable objects. | Step‑by‑step guide to add sample data, build a PivotTable, activate ribbon tools, and save the workbook in a .NET console app.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRibbonDemo
{
    // Creates a workbook, inserts sample sales data, builds a PivotTable on a separate sheet, assigns row and data fields, and turns on the PivotTable UI options—field list pane, field dialog, and wizard—before refreshing, calculating, and saving the file as PivotTable_With_Ribbon.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (will hold source data)
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Product");
            dataSheet.Cells["C1"].PutValue("Sales");

            dataSheet.Cells["A2"].PutValue("Electronics");
            dataSheet.Cells["B2"].PutValue("Laptop");
            dataSheet.Cells["C2"].PutValue(1200);

            dataSheet.Cells["A3"].PutValue("Electronics");
            dataSheet.Cells["B3"].PutValue("Phone");
            dataSheet.Cells["C3"].PutValue(800);

            dataSheet.Cells["A4"].PutValue("Furniture");
            dataSheet.Cells["B4"].PutValue("Chair");
            dataSheet.Cells["C4"].PutValue(150);

            dataSheet.Cells["A5"].PutValue("Furniture");
            dataSheet.Cells["B5"].PutValue("Table");
            dataSheet.Cells["C5"].PutValue(300);

            // Add a new worksheet that will contain the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create the pivot table using the source range from the data sheet
            // The source range is specified with an external reference to the Data sheet
            int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C5", "A3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add fields to the pivot table areas
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Enable the built‑in PivotTable ribbon features:
            // - Field List (shows the pane on the right)
            // - Field Dialog (double‑click field dialog)
            // - Wizard (PivotTable Wizard)
            pivotTable.EnableFieldList = true;
            pivotTable.EnableFieldDialog = true;
            pivotTable.EnableWizard = true;

            // Refresh and calculate the pivot data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_With_Ribbon.xlsx");
        }
    }
}
