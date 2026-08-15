// Title: Add a Profit‑Margin Calculated Field to an Aspose.Cells Pivot Table in C# (.NET)
// Description: Creates a workbook, inserts Product, Revenue, and Cost data, builds a pivot table, adds a calculated field named ProfitMargin using the formula =(Revenue‑Cost)/Revenue, formats it as a percentage, refreshes the pivot cache, recalculates the table, and saves the result as an XLSX file.
// Keywords: Aspose.Cells | C# pivot table | calculated field | profit margin formula | Excel financial report | .NET Excel automation | pivot cache refresh | percentage formatting | Aspose.Cells example | AddCalculatedField
// Common Searches: how to add a calculated field to an Aspose.Cells pivot table | profit margin formula Aspose.Cells C# | format pivot table field as percentage Aspose.Cells | refresh pivot cache after adding calculated field .NET | Aspose.Cells pivot table example with revenue and cost
// Developer Intent: Generate a pivot table and insert a calculated field that derives profit margin from Revenue and Cost columns.
// Use Cases: Produce a sales dashboard that shows profit margin per product directly in the pivot view. | Create an Excel workbook for financial analysis where margin percentages are calculated on‑the‑fly. | Export a ready‑to‑present report with profit‑margin calculations embedded in the pivot table.
// AI Prompts: Write C# code using Aspose.Cells to add a calculated field called ProfitMargin to a pivot table with the formula (Revenue‑Cost)/Revenue and format it as a percentage. | Explain the steps to refresh the pivot cache and recalculate data after inserting a calculated field in Aspose.Cells for .NET. | Show how to save the workbook to a custom directory after adding a profit‑margin calculated field to a pivot table.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Creates a workbook, inserts Product, Revenue, and Cost data, builds a pivot table, adds a calculated field named ProfitMargin using the formula =(Revenue‑Cost)/Revenue, formats it as a percentage, refreshes the pivot cache, recalculates the table, and saves the result as an XLSX file.
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

                // Populate sample data with Revenue and Cost columns
                cells["A1"].Value = "Product";
                cells["B1"].Value = "Revenue";
                cells["C1"].Value = "Cost";

                cells["A2"].Value = "A";
                cells["B2"].Value = 1200;
                cells["C2"].Value = 800;

                cells["A3"].Value = "B";
                cells["B3"].Value = 1500;
                cells["C3"].Value = 900;

                cells["A4"].Value = "C";
                cells["B4"].Value = 2000;
                cells["C4"].Value = 1300;

                // Create a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Revenue");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Cost");

                // Add a calculated field for profit margin: (Revenue - Cost) / Revenue
                // The formula must start with '=' and reference the source field names
                pivotTable.AddCalculatedField("ProfitMargin", "=(Revenue-Cost)/Revenue", true);

                // Optionally format the calculated field as a percentage
                PivotField profitMarginField = pivotTable.DataFields[pivotTable.DataFields.Count - 1];
                profitMarginField.NumberFormat = "0.00%";

                // Refresh the pivot cache and calculate the pivot table
                pivotTable.RefreshData();      // Correct API to refresh cache
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotTable_With_ProfitMargin.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

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
            AddProfitMarginCalculatedField.Run();
        }
    }
}
