// Title: How to apply a custom currency number format and bold style to a PivotTable grand total row using Aspose.Cells for C#
// AI Prompts: Generate C# code with Aspose.Cells that sets a currency number format (e.g., $#,##0.00) for the PivotTable’s total line. | Show how to locate the Grand Total label cell in a PivotTable and apply a bold, dark‑blue font style to the entire line using Aspose.Cells. | Explain the steps to refresh and recalculate a PivotTable after modifying its number format and row style with Aspose.Cells in .NET.
// Common Searches: Aspose.Cells C# apply custom currency pattern to pivot table totals | make pivot table overall total row bold dark blue with Aspose.Cells | update pivot table data after applying new number format using Aspose.Cells .NET
// Tags: pivot table total line custom number format Aspose.Cells | apply highlighted styling to pivot total line C# | recalculate pivot calculations following number format update Aspose.Cells | set data field number format for pivot table Aspose.Cells | style grand total line in Excel workbook using Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable, PivotField, PivotFieldType

namespace AsposeCellsExamples
{
    // The example creates a workbook, fills it with sample sales data, adds a PivotTable, assigns a currency number format to the data field, defines a style with bold dark‑blue font, finds the Grand Total label, applies the style to the entire total line, refreshes and recalculates the PivotTable, and saves the file as PivotTable_GrandTotal_CustomFormat.xlsx.
    class Program
    {
        static void Main(string[] args)
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

                // Configure the pivot table fields
                pivot.AddFieldToArea(PivotFieldType.Row, "Product");
                pivot.AddFieldToArea(PivotFieldType.Column, "Region");
                int dataFieldIdx = pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Set the number format for the data field (e.g., currency)
                PivotField dataField = pivot.DataFields[dataFieldIdx];
                dataField.NumberFormat = "$#,##0.00";

                // Refresh and calculate the pivot table to populate data
                pivot.RefreshData();
                pivot.CalculateData();

                // Create a style for the grand total row
                Style grandTotalStyle = workbook.CreateStyle();
                grandTotalStyle.Custom = "$#,##0.00";
                grandTotalStyle.Font.IsBold = true;
                grandTotalStyle.Font.Color = Color.DarkBlue;

                // Locate the cell that contains the Grand Total label
                Cell grandTotalLabelCell = sheet.Cells.Find(pivot.GrandTotalName, null);
                if (grandTotalLabelCell != null)
                {
                    // Apply the style to the entire grand total row
                    pivot.FormatRow(grandTotalLabelCell.Row, grandTotalStyle);
                }

                // Save the workbook
                string outputPath = "PivotTable_GrandTotal_CustomFormat.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
