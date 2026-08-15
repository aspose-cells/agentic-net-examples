// Title: Set a custom tooltip (Alt Text) for an Aspose.Cells PivotTable in C#
// Description: This C# example shows how to create a workbook, fill it with sample sales data, add a PivotTable, enable data tips, and assign custom AltTextDescription and AltTextTitle values so a tooltip appears when the user hovers over the PivotTable. The file is saved as PivotTableWithCustomTooltip.xlsx.
// Keywords: Aspose.Cells | C# | PivotTable | AltTextDescription | AltTextTitle | ShowDataTips | custom tooltip | Excel tooltip | pivot table tooltip .NET | Aspose.Cells tooltip example
// Common Searches: Aspose.Cells set pivot table tooltip C# | How to add AltTextDescription to PivotTable using Aspose.Cells | Enable ShowDataTips for PivotTable Aspose .NET | Add hover text to Excel PivotTable programmatically | Aspose.Cells PivotTable AltTextTitle example
// Developer Intent: Programmatically add a hover tooltip to a PivotTable by setting AltTextDescription (and optionally AltTextTitle) and turning on ShowDataTips.
// Use Cases: Provide explanatory hover text for end‑users viewing the generated Excel report | Add contextual information to automated reporting workbooks | Improve accessibility by supplying alt text for screen readers | Standardize tooltip content across multiple PivotTables created by code
// AI Prompts: Generate C# code with Aspose.Cells that creates a PivotTable and sets a custom tooltip using AltTextDescription and AltTextTitle. | Explain how ShowDataTips, AltTextDescription, and AltTextTitle work together to display a tooltip for a PivotTable in Aspose.Cells. | Show how to assign different tooltips to individual PivotTable fields using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotTooltipDemo
{
    // This C# example shows how to create a workbook, fill it with sample sales data, add a PivotTable, enable data tips, and assign custom AltTextDescription and AltTextTitle values so a tooltip appears when the user hovers over the PivotTable. The file is saved as PivotTableWithCustomTooltip.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Region");
            worksheet.Cells["C1"].PutValue("Sales");

            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue("North");
            worksheet.Cells["C2"].PutValue(1200);

            worksheet.Cells["A3"].PutValue("Apple");
            worksheet.Cells["B3"].PutValue("South");
            worksheet.Cells["C3"].PutValue(800);

            worksheet.Cells["A4"].PutValue("Banana");
            worksheet.Cells["B4"].PutValue("North");
            worksheet.Cells["C4"].PutValue(1500);

            worksheet.Cells["A5"].PutValue("Banana");
            worksheet.Cells["B5"].PutValue("South");
            worksheet.Cells["C5"].PutValue(1100);

            // Add a pivot table based on the data range
            PivotTableCollection pivotTables = worksheet.PivotTables;
            int pivotIndex = pivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Enable data tips so the tooltip will be shown
            pivotTable.ShowDataTips = true;

            // Set a custom tooltip (alt text description) for the entire pivot table
            // This description appears as a tooltip when the user hovers over the pivot table
            pivotTable.AltTextDescription = "Sales breakdown by product and region";

            // Optionally set a title for the alt text
            pivotTable.AltTextTitle = "Sales Pivot Table";

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTableWithCustomTooltip.xlsx");
        }
    }
}
