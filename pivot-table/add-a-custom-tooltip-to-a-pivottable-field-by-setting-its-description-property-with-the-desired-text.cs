// Title: Set a custom tooltip (Alt Text) for an Aspose.Cells PivotTable in C#
// Description: This example creates a workbook, adds sample sales data, builds a PivotTable, and assigns a custom tooltip by setting the PivotTable.AltTextDescription and optionally PivotTable.AltTextTitle properties. The pivot is refreshed, calculated, and saved as an Excel file, demonstrating how to provide hover text and accessibility notes for end users.
// Keywords: Aspose.Cells PivotTable tooltip | C# AltTextDescription | AltTextTitle Aspose.Cells | custom pivot tooltip .NET | Excel accessibility Alt Text | Aspose.Cells set pivot description
// Common Searches: how to add tooltip to Aspose.Cells PivotTable C# | set AltTextDescription for PivotTable in .NET | Aspose.Cells pivot alt text title example | add custom hover text to Excel pivot using Aspose | accessibility description for Aspose pivot table
// Developer Intent: Apply a custom tooltip to a PivotTable by assigning text to its AltTextDescription (and optionally AltTextTitle) properties in Aspose.Cells for .NET.
// Use Cases: Show explanatory hover text when users mouse over the pivot in Excel. | Provide screen‑reader friendly descriptions for accessibility compliance. | Summarize the purpose of a shared pivot table without opening the workbook.
// AI Prompts: Generate C# code that creates a PivotTable with a custom AltTextDescription using Aspose.Cells. | Explain the effect of AltTextDescription and AltTextTitle on Excel tooltip behavior. | Show how to add accessibility alt text to a PivotTable field in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotTooltipDemo
{
    // This example creates a workbook, adds sample sales data, builds a PivotTable, and assigns a custom tooltip by setting the PivotTable.AltTextDescription and optionally PivotTable.AltTextTitle properties. The pivot is refreshed, calculated, and saved as an Excel file, demonstrating how to provide hover text and accessibility notes for end users.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Clothing");
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["A4"].PutValue("Electronics");
            sheet.Cells["B4"].PutValue(1500);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Set a custom tooltip (alt text description) for the pivot table
            pivotTable.AltTextDescription = "This pivot shows sales amount by category.";

            // Optionally set a title for the alt text
            pivotTable.AltTextTitle = "Sales Pivot Tooltip";

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableWithCustomTooltip.xlsx");
        }
    }
}
