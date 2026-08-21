// Title: ShowErrorValues in Aspose.Cells PivotTable – Display #N/A for Null Data (C#)
// Description: Creates a workbook, adds sample rows with null values, builds a PivotTable, assigns Category to rows and Value to data, enables DisplayErrorString, sets the error text to "#N/A", refreshes the pivot, and saves the file. Demonstrates how to make missing entries appear as #N/A in the generated Excel pivot.
// Keywords: Aspose.Cells | PivotTable | ShowErrorValues | DisplayErrorString | #N/A | null values | .NET | C# | custom error string | missing data | Excel export
// Common Searches: Aspose.Cells show error values in pivot table | Display #N/A for null entries using Aspose.Cells | Enable DisplayErrorString in C# PivotTable | Set custom error text for missing data Aspose.Cells | PivotTable error handling .NET
// Developer Intent: Activate ShowErrorValues so null cells are shown as #N/A in the pivot table.
// Use Cases: Financial statements where blank sales figures are highlighted as #N/A for audit clarity. | Data‑quality dashboards that flag unavailable values directly in Excel pivots. | Automated report generation that uses a custom error indicator to guide downstream analysts.
// AI Prompts: Generate C# code with Aspose.Cells that sets DisplayErrorString = true and ErrorString = "#N/A" on a PivotTable. | Explain the steps to refresh and calculate a PivotTable after enabling ShowErrorValues in Aspose.Cells. | Show how to modify an existing workbook to replace null values with a custom error string in its pivot view.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotShowErrorValuesDemo
{
    // Creates a workbook, adds sample rows with null values, builds a PivotTable, assigns Category to rows and Value to data, enables DisplayErrorString, sets the error text to "#N/A", refreshes the pivot, and saves the file. Demonstrates how to make missing entries appear as #N/A in the generated Excel pivot.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with some missing (null) values
            // A column: Category, B column: Value
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(null); // missing data point
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(300);
            sheet.Cells["A5"].PutValue("D");
            sheet.Cells["B5"].PutValue(null); // another missing data point

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot fields (Category as row, Value as data)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column A (Category)
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column B (Value)

            // Enable display of custom error string for cells that contain errors/missing data
            pivotTable.DisplayErrorString = true;
            // Set the custom error string to "#N/A" to clearly indicate missing values
            pivotTable.ErrorString = "#N/A";

            // Refresh data and calculate the pivot table to apply the settings
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the configured pivot table
            workbook.Save("PivotTable_ShowErrorValues.xlsx");

            Console.WriteLine("Pivot table created with ShowErrorValues enabled (displaying #N/A for missing data).");
        }
    }
}
