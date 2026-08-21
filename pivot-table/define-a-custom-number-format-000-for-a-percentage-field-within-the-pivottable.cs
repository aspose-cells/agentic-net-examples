// Title: Aspose.Cells .NET – Set PivotTable data field to custom 0.00% format (C#)
// Description: Creates a workbook, adds Category and Percent columns, builds a PivotTable, places Category in rows and Percent in the data area, then applies the custom number format "0.00%" to the data field before refreshing, calculating, and saving the file.
// Keywords: Aspose.Cells PivotTable number format | C# custom percentage format | 0.00% pivot data field | Aspose.Cells set NumberFormat | Excel pivot percentage formatting
// Common Searches: Aspose.Cells set pivot data field format C# | custom 0.00% number format in PivotTable .NET | how to display percentages with two decimals in Aspose.Cells | change pivot table number format Aspose.Cells
// Developer Intent: Apply the custom number format "0.00%" to a PivotTable data field using Aspose.Cells for .NET.
// Use Cases: Show sales commission rates with two decimal places in automated reports. | Display conversion or click‑through rates as percentages in financial dashboards. | Standardize inventory shrinkage percentages across exported Excel files.
// AI Prompts: Write C# code with Aspose.Cells that creates a PivotTable and sets the data field's NumberFormat to "0.00%". | Explain how to modify an existing Aspose.Cells PivotTable to use a percentage format with two decimal places for a data field. | Provide a step‑by‑step guide for applying a custom 0.00% number format to a PivotTable data field in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableCustomNumberFormat
{
    // Creates a workbook, adds Category and Percent columns, builds a PivotTable, places Category in rows and Percent in the data area, then applies the custom number format "0.00%" to the data field before refreshing, calculating, and saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (Category and Percentage values)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Percent");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(0.1234); // 12.34%
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(0.5678); // 56.78%
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(0.9012); // 90.12%

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PercentPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add the category field to the row area
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add the percent field to the data area
            int dataFieldPos = pivot.AddFieldToArea(PivotFieldType.Data, "Percent");
            PivotField dataField = pivot.DataFields[dataFieldPos];

            // Set a custom number format to display values as percentages with two decimals
            dataField.NumberFormat = "0.00%";

            // Refresh and calculate the pivot table to apply changes
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableCustomNumberFormat.xlsx");
        }
    }
}
