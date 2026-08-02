// Title: Aspose.Cells for .NET – Apply Currency Formatting to PivotTable Grand Total (C#)
// Description: Creates a workbook, adds sample sales data, builds a PivotTable, assigns the Sales field as a data field, applies the "$#,##0.00" format to the data field (affecting all cells and the grand total row), refreshes the pivot, and saves the Excel file.
// Keywords: Aspose.Cells | C# PivotTable number format | currency format pivot total | grand total formatting Aspose | custom number format Excel | PivotTable data field format | Aspose.Cells .NET example
// Common Searches: Aspose.Cells set number format for pivot table total | C# format grand total row in Excel pivot | apply currency format to pivot data field Aspose | how to format pivot table totals using Aspose.Cells | custom number format for pivot table in .NET
// Developer Intent: Set a currency number format for the PivotTable’s grand total row.
// Use Cases: Financial reports where totals must appear as $ values | Exporting sales analysis to Excel with consistent monetary styling | Automated dashboard generation that requires formatted grand totals | Preparing audit‑ready spreadsheets with standardized currency display
// AI Prompts: Write C# code with Aspose.Cells that formats the grand total row of a PivotTable using a custom currency pattern. | Show how to apply a percentage number format only to the grand total row of a PivotTable in Aspose.Cells. | Explain the steps to change the number format of a PivotTable data field without affecting the underlying source data in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable, PivotField, PivotFieldType

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample sales data, builds a PivotTable, assigns the Sales field as a data field, applies the "$#,##0.00" format to the data field (affecting all cells and the grand total row), refreshes the pivot, and saves the Excel file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Region");
                worksheet.Cells["C1"].PutValue("Sales");

                worksheet.Cells["A2"].PutValue("Laptop");
                worksheet.Cells["B2"].PutValue("North");
                worksheet.Cells["C2"].PutValue(1200);

                worksheet.Cells["A3"].PutValue("Laptop");
                worksheet.Cells["B3"].PutValue("South");
                worksheet.Cells["C3"].PutValue(1500);

                worksheet.Cells["A4"].PutValue("Phone");
                worksheet.Cells["B4"].PutValue("North");
                worksheet.Cells["C4"].PutValue(800);

                worksheet.Cells["A5"].PutValue("Phone");
                worksheet.Cells["B5"].PutValue("South");
                worksheet.Cells["C5"].PutValue(1100);

                // Add a pivot table based on the data range
                int pivotIndex = worksheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
                int dataFieldPos = pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Retrieve the data field that will hold the sales values
                PivotField dataField = pivotTable.DataFields[dataFieldPos];

                // Set a custom number format – this format is applied to all data cells,
                // including the grand total row
                dataField.NumberFormat = "$#,##0.00";

                // Refresh and calculate the pivot table to apply the format
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Ensure the output directory exists
                string outputPath = "PivotTableGrandTotalNumberFormat.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the pivot table:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
