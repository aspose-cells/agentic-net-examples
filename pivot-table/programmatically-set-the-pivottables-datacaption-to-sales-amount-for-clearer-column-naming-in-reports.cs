// Title: Aspose.Cells .NET: Set PivotTable DataFieldHeaderName to "Sales Amount"
// Description: Learn how to create a workbook with sample sales data, add a PivotTable, and programmatically rename the data field header to "Sales Amount" using the DataFieldHeaderName property in Aspose.Cells for C#. The example refreshes and calculates the PivotTable before saving the file.
// Keywords: Aspose.Cells | .NET | C# | PivotTable | DataFieldHeaderName | custom data caption | Excel automation | sales report pivot | programmatic Excel | global developers
// Common Searches: Aspose.Cells set PivotTable data caption C# | DataFieldHeaderName property example | rename PivotTable value column Aspose | change pivot table header Aspose.Cells .NET | customize PivotTable column title programmatically
// Developer Intent: Rename the PivotTable's data column header to a more descriptive label such as "Sales Amount" for clearer reporting.
// Use Cases: Generate sales dashboards where the values column reads "Sales Amount" instead of the raw field name. | Create multiple PivotTables with metric‑specific captions like "Revenue" or "Units Sold" in automated reports. | Standardize Excel exports for finance teams that require explicit column headings across different regions.
// AI Prompts: Show C# code using Aspose.Cells to set a PivotTable's DataFieldHeaderName to a custom string and refresh the table. | How can I rename the data field header of an existing PivotTable in a workbook loaded with Aspose.Cells? | Explain the difference between DataFieldHeaderName and other caption properties in Aspose.Cells and give usage examples.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableDataCaptionExample
{
    // Learn how to create a workbook with sample sales data, add a PivotTable, and programmatically rename the data field header to "Sales Amount" using the DataFieldHeaderName property in Aspose.Cells for C#. The example refreshes and calculates the PivotTable before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            Cells cells = sheet.Cells;
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Region";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "Bike";
            cells["B2"].Value = "North";
            cells["C2"].Value = 1200;

            cells["A3"].Value = "Bike";
            cells["B3"].Value = "South";
            cells["C3"].Value = 1500;

            cells["A4"].Value = "Car";
            cells["B4"].Value = "North";
            cells["C4"].Value = 2000;

            cells["A5"].Value = "Car";
            cells["B5"].Value = "South";
            cells["C5"].Value = 2500;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);    // Product as row field
            pivotTable.AddFieldToArea(PivotFieldType.Column, 1); // Region as column field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 2);   // Sales as data field

            // Set the data field header caption to "Sales Amount"
            pivotTable.DataFieldHeaderName = "Sales Amount";

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableWithDataCaption.xlsx");
        }
    }
}
