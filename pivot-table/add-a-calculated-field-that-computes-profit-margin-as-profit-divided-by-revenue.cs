// Title: Add a Profit Margin Calculated Field to a Pivot Table with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, inserts product, revenue and profit data, builds a pivot table, adds a calculated field ProfitMargin = Profit/Revenue, formats it as a percentage, refreshes the pivot and saves the file.
// Keywords: Aspose.Cells | C# | pivot table calculated field | profit margin | percentage format | RefreshData | CalculateData | Excel automation | financial reporting | GitHub Aspose.Cells example
// Common Searches: Aspose.Cells add calculated field profit margin | C# pivot table percentage format Aspose.Cells | Refresh pivot table after calculated field Aspose.Cells | How to compute profit margin in Aspose.Cells pivot | Aspose.Cells pivot table example GitHub
// Developer Intent: Programmatically add a ProfitMargin calculated field to an Aspose.Cells pivot table and display it as a percentage.
// Use Cases: Build a sales dashboard that shows revenue, profit and profit‑margin percentages per product. | Automate quarterly financial reports by injecting a profit‑margin field into existing pivot tables across many workbooks. | Prepare Excel files for BI tools where profit margin must be pre‑calculated and formatted as a percentage.
// AI Prompts: Generate C# code using Aspose.Cells to add a calculated field named ProfitMargin that divides Profit by Revenue and format it as a percentage. | Explain how to refresh and recalculate a pivot table after adding a calculated field in Aspose.Cells for .NET. | Provide a concise Aspose.Cells example that creates a workbook, adds sample data, builds a pivot table, inserts a profit‑margin calculated field, applies percentage formatting, and saves the file.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Creates a workbook, inserts product, revenue and profit data, builds a pivot table, adds a calculated field ProfitMargin = Profit/Revenue, formats it as a percentage, refreshes the pivot and saves the file.
    public class AddProfitMarginCalculatedField
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data: Product, Revenue, Profit
                cells["A1"].Value = "Product";
                cells["B1"].Value = "Revenue";
                cells["C1"].Value = "Profit";

                cells["A2"].Value = "A";
                cells["B2"].Value = 1000;
                cells["C2"].Value = 200;

                cells["A3"].Value = "B";
                cells["B3"].Value = 1500;
                cells["C3"].Value = 300;

                cells["A4"].Value = "C";
                cells["B4"].Value = 2000;
                cells["C4"].Value = 500;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");   // Row field
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Revenue"); // Data field
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Profit");  // Data field

                // Add a calculated field "ProfitMargin" = Profit / Revenue and drag it to the data area
                pivotTable.AddCalculatedField("ProfitMargin", "=Profit/Revenue", true);

                // Optionally format the calculated field as percentage
                PivotField profitMarginField = pivotTable.DataFields[pivotTable.DataFields.Count - 1];
                profitMarginField.NumberFormat = "0.00%";

                // Refresh and calculate the pivot table using the correct API
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotTable_With_ProfitMargin.xlsx");
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
            AddProfitMarginCalculatedField.Run();
        }
    }
}
