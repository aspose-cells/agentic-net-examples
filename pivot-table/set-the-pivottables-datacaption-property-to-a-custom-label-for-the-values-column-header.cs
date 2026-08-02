// Title: Set a custom values column header (DataFieldHeaderName) for a PivotTable with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, populates sample sales data, adds a PivotTable, places Product in rows and Sales in the values area, then assigns a custom label to the values column header by setting PivotTable.DataFieldHeaderName, refreshes the pivot, and saves the file.
// Keywords: Aspose.Cells | C# PivotTable | DataFieldHeaderName | custom values column header | set pivot table data caption | Aspose.Cells for .NET example | change pivot values header | Aspose.Cells GitHub example | SetPivotTableDataCaptionDemo
// Common Searches: Aspose.Cells set PivotTable DataFieldHeaderName | change values column header in Aspose.Cells PivotTable | custom data caption for PivotTable C# | Aspose.Cells PivotTable column header label example | how to rename values field in Aspose.Cells pivot
// Developer Intent: Assign a custom label to the values column header of a PivotTable programmatically.
// Use Cases: Generate a sales report where the values column reads "Total Sales" instead of the default "Sum of Sales". | Create a financial summary workbook with a data field header like "Amount (USD)" to match corporate terminology. | Build a product performance sheet that displays "Revenue" as the pivot data caption while preserving calculations.
// AI Prompts: Show how to set PivotTable.DataFieldHeaderName in Aspose.Cells for .NET and refresh the pivot. | Provide a C# code snippet that changes the values column header of an existing PivotTable to a custom string using Aspose.Cells. | Explain the steps to customize a PivotTable data caption and ensure the pivot recalculates correctly.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Creates a workbook, populates sample sales data, adds a PivotTable, places Product in rows and Sales in the values area, then assigns a custom label to the values column header by setting PivotTable.DataFieldHeaderName, refreshes the pivot, and saves the file.
    public class SetPivotTableDataCaptionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Region");
                sheet.Cells["C1"].PutValue("Sales");

                sheet.Cells["A2"].PutValue("Laptop");
                sheet.Cells["B2"].PutValue("North");
                sheet.Cells["C2"].PutValue(1000);

                sheet.Cells["A3"].PutValue("Laptop");
                sheet.Cells["B3"].PutValue("South");
                sheet.Cells["C3"].PutValue(1500);

                sheet.Cells["A4"].PutValue("Phone");
                sheet.Cells["B4"].PutValue("North");
                sheet.Cells["C4"].PutValue(800);

                sheet.Cells["A5"].PutValue("Phone");
                sheet.Cells["B5"].PutValue("South");
                sheet.Cells["C5"].PutValue(1200);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: rows = Product, values = Sales
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column
                pivotTable.AddFieldToArea(PivotFieldType.Data, 2);  // Sales column

                // Set a custom caption for the values column header (DataFieldHeaderName)
                pivotTable.DataFieldHeaderName = "Custom Values";

                // Refresh and calculate the pivot table to apply changes
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("SetPivotTableDataCaptionDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetPivotTableDataCaptionDemo.Run();
        }
    }
}
