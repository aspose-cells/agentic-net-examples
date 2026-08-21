// Title: Apply a Custom Currency Number Format to PivotTable Grand Total Row with Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, adds sample sales data, builds a PivotTable, sets the data field to sum, applies the number format "$#,##0.00" to the data field (which also formats the row grand total), enables row grand totals, refreshes the PivotTable, and saves the file as an .xlsx.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | custom number format | grand total formatting | currency format | Excel export | data field formatting | financial reporting
// Common Searches: Aspose.Cells format pivot table grand total | C# set number format for pivot table totals Aspose | How to apply currency format to PivotTable row totals using Aspose.Cells | Custom number format for PivotTable data field .NET | Show row grand totals with formatting Aspose.Cells
// Developer Intent: Programmatically set a custom number format for a PivotTable data field so that both data cells and the row grand total display values in the specified format using Aspose.Cells for .NET.
// Use Cases: Generate sales reports with USD currency totals in the grand‑total row. | Create financial statements where pivot totals use an accounting number format. | Export Excel files with consistent formatting for data cells and grand totals. | Automate dashboards that require formatted row grand totals. | Prepare pivot‑based summaries for multinational data with locale‑specific number formats.
// AI Prompts: Write C# code using Aspose.Cells to create a PivotTable and apply a custom currency number format to its row grand total. | Show how to set a percentage number format with two decimal places for a PivotTable data field and its grand total in Aspose.Cells for .NET. | Provide an example that enables row grand totals and applies an accounting format (e.g., "_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)" ) to a PivotTable using Aspose.Cells C#.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, adds sample sales data, builds a PivotTable, sets the data field to sum, applies the number format "$#,##0.00" to the data field (which also formats the row grand total), enables row grand totals, refreshes the PivotTable, and saves the file as an .xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Region");
                sheet.Cells["C1"].PutValue("Sales");

                sheet.Cells["A2"].PutValue("Laptop");
                sheet.Cells["B2"].PutValue("North");
                sheet.Cells["C2"].PutValue(1200);

                sheet.Cells["A3"].PutValue("Laptop");
                sheet.Cells["B3"].PutValue("South");
                sheet.Cells["C3"].PutValue(1500);

                sheet.Cells["A4"].PutValue("Phone");
                sheet.Cells["B4"].PutValue("North");
                sheet.Cells["C4"].PutValue(800);

                sheet.Cells["A5"].PutValue("Phone");
                sheet.Cells["B5"].PutValue("South");
                sheet.Cells["C5"].PutValue(1100);

                // Add a pivot table based on the data range
                int ptIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
                PivotTable pivot = sheet.PivotTables[ptIndex];

                // Configure pivot fields
                pivot.AddFieldToArea(PivotFieldType.Row, "Product");
                int dataFieldIdx = pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Set custom number format for the data field (applies to data cells and grand total row)
                PivotField dataField = pivot.DataFields[dataFieldIdx];
                dataField.Function = ConsolidationFunction.Sum;
                dataField.NumberFormat = "$#,##0.00";

                // Ensure the grand total row is displayed
                pivot.ShowRowGrandTotals = true;

                // Refresh and calculate the pivot table
                pivot.RefreshData();
                pivot.CalculateData();

                // Define output file path
                string outputPath = "PivotTableGrandTotalCustomFormat.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
