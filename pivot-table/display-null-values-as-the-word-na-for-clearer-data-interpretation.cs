// Title: Display "N/A" for null cells in an Aspose.Cells pivot table using C#
// AI Prompts: Create a C# program that builds a workbook, adds rows with null entries, generates a pivot table, and configures the pivot to show "N/A" for null cells by setting DisplayNullString and NullString. | Show how to enable custom null string handling for a pivot table in Aspose.Cells .NET and save the workbook as an XLSX file.
// Common Searches: Aspose.Cells pivot table replace blank values with N/A using C# | How to configure null value display in an Aspose.Cells generated pivot | Using DisplayNullString to show N/A in Aspose.Cells pivot tables | C# code sample for null string handling in Aspose.Cells pivot | Aspose.Cells C# pivot table custom null placeholder example
// Tags: Aspose.Cells pivot null placeholder | C# DisplayNullString Aspose.Cells | Aspose.Cells pivot custom null text | Excel pivot N/A display Aspose | Aspose.Cells workbook null handling

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsNullDisplayDemo
{
    // The example creates a workbook, inserts sample data containing null entries, adds a pivot table, enables custom null display by setting DisplayNullString to true and NullString to "N/A", refreshes and calculates the pivot, and saves the file as PivotTable_NullAsNA.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with null values
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(1500);
            sheet.Cells["A4"].PutValue(null);      // Null product name
            sheet.Cells["B4"].PutValue(null);      // Null sales value

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: rows = Product, data = Sales
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

            // Enable custom display for null values and set the desired string
            pivotTable.DisplayNullString = true;
            pivotTable.NullString = "N/A";

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_NullAsNA.xlsx");
        }
    }
}
