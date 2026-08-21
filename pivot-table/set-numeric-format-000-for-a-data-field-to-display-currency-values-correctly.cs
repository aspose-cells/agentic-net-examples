// Title: Aspose.Cells for .NET – Apply '#,##0.00' number format to a PivotTable data field
// Description: Demonstrates how to create a workbook, add sample product‑sales data, build a PivotTable, set the Sales field to Sum, and assign the currency format '#,##0.00' to the data field. The example refreshes the pivot cache, calculates the results, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells PivotTable number format | set currency format pivot .NET | custom numeric format '#,##0.00' Aspose | format pivot data field Aspose.Cells | PivotField NumberFormat property | Excel pivot table formatting C# | Aspose.Cells example currency formatting
// Common Searches: Aspose.Cells set number format for pivot table | how to format pivot data as currency in .NET | apply '#,##0.00' to PivotField Aspose.Cells | pivot table number format example C# | change numeric format of pivot table values
// Developer Intent: Apply a currency‑style numeric format '#,##0.00' to a PivotTable data field using Aspose.Cells for .NET.
// Use Cases: Generate financial reports where summed sales appear with thousand separators and two decimals. | Create dashboards that export Excel files with pivot values displayed as currency. | Standardize numeric appearance across multiple pivot tables in an automated workbook generation process.
// AI Prompts: Show how to set the NumberFormat property of a PivotField to '#,##0.00' in Aspose.Cells for C#. | Provide a code example that formats pivot table sum values as currency and refreshes the cache. | Explain the steps to change a pivot table data field’s numeric format after defining its aggregation function.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to create a workbook, add sample product‑sales data, build a PivotTable, set the Sales field to Sum, and assign the currency format '#,##0.00' to the data field. The example refreshes the pivot cache, calculates the results, and saves the file as an Excel workbook.
public class SetNumberFormatDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(1234.56);
            worksheet.Cells["A3"].PutValue("Orange");
            worksheet.Cells["B3"].PutValue(2345.78);
            worksheet.Cells["A4"].PutValue("Banana");
            worksheet.Cells["B4"].PutValue(3456.90);

            // Add a pivot table based on the data range
            int pivotIndex = worksheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Add the product field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

            // Add the sales field to the data area and obtain the PivotField object
            int dataFieldIndex = pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
            PivotField dataField = pivotTable.DataFields[dataFieldIndex];
            dataField.Function = ConsolidationFunction.Sum;

            // Set the custom numeric format '#,##0.00' for the data field
            dataField.NumberFormat = "#,##0.00";

            // Refresh pivot cache and calculate data to apply the format
            pivotTable.RefreshData();      // Correct method to refresh cache
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotNumberFormatDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}

public class Program
{
    public static void Main()
    {
        SetNumberFormatDemo.Run();
    }
}
