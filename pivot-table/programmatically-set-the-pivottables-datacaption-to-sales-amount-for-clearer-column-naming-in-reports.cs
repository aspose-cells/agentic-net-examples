// Title: How to Set a Custom Data Caption (DataFieldHeaderName) for an Aspose.Cells PivotTable in C#
// Description: Demonstrates creating a workbook with sample sales data, adding a PivotTable, assigning row/column/data fields, and using the DataFieldHeaderName property to rename the data area header to "Sales Amount" before refreshing, calculating, and saving the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells PivotTable DataFieldHeaderName | C# set pivot table data caption | custom data field header Aspose.Cells | rename pivot table data area header | Aspose.Cells .NET example | Excel PivotTable custom caption | pivot table column naming Aspose | Aspose.Cells sample code GitHub | programmatic Excel report C# | pivot table data caption API
// Common Searches: Aspose.Cells how to change data field header in pivot table C# | set DataFieldHeaderName property Aspose.Cells | rename pivot table data caption programmatically | C# example for custom pivot table header Aspose | Aspose.Cells pivot table column naming tutorial
// Developer Intent: Apply a custom label to the PivotTable data area header (e.g., "Sales Amount") using Aspose.Cells for .NET.
// Use Cases: Produce sales reports where the data column shows a friendly name instead of the source column title. | Standardize metric headings across multiple generated PivotTables (Revenue, Quantity, Profit). | Automate Excel workbook creation for business users who require clear, self‑describing column headings.
// AI Prompts: Generate C# code with Aspose.Cells that sets the DataFieldHeaderName of a PivotTable to a custom string. | Show how to modify the data caption of an existing Aspose.Cells PivotTable after it has been created. | Explain the role of DataFieldHeaderName, RefreshData, and CalculateData when customizing PivotTable headers in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // Demonstrates creating a workbook with sample sales data, adding a PivotTable, assigning row/column/data fields, and using the DataFieldHeaderName property to rename the data area header to "Sales Amount" before refreshing, calculating, and saving the file with Aspose.Cells for .NET.
    public class SetPivotDataCaption
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Product";
                sheet.Cells["B1"].Value = "Region";
                sheet.Cells["C1"].Value = "Sales";

                sheet.Cells["A2"].Value = "Bike";
                sheet.Cells["B2"].Value = "North";
                sheet.Cells["C2"].Value = 1200;

                sheet.Cells["A3"].Value = "Bike";
                sheet.Cells["B3"].Value = "South";
                sheet.Cells["C3"].Value = 1500;

                sheet.Cells["A4"].Value = "Car";
                sheet.Cells["B4"].Value = "North";
                sheet.Cells["C4"].Value = 2000;

                sheet.Cells["A5"].Value = "Car";
                sheet.Cells["B5"].Value = "South";
                sheet.Cells["C5"].Value = 2500;

                // Add a pivot table to the worksheet (source range A1:C5, destination E3)
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Product as row field
                pivotTable.AddFieldToArea(PivotFieldType.Column, 1);   // Region as column field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 2);     // Sales as data field

                // Set a custom caption for the data (value) area header
                pivotTable.DataFieldHeaderName = "Sales Amount";

                // Refresh and calculate the pivot table to apply changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotTableWithDataCaption.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetPivotDataCaption.Run();
        }
    }
}
