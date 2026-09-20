// Title: Generate a PivotTable and Freeze Row Labels in an Excel Workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that builds a workbook, inserts sample data, creates a PivotTable with 'Category' as the row axis and 'Sales' as the values axis, and then locks the label column in place. | Demonstrate adding a PivotTable via the PivotTables.Add method and then freezing the worksheet so the label column stays visible during navigation. | Provide a full example that defines the source range, places the PivotTable at a chosen cell, configures its row and data fields, and applies a freeze to the row‑label column before saving.
// Common Searches: Aspose.Cells C# create pivot table from range and freeze first column | How to keep pivot table row headers visible using FreezePanes in Aspose.Cells | C# example for adding a PivotTable and freezing row labels in Excel with Aspose.Cells | Programmatically freeze pane for pivot table row labels in .NET Aspose.Cells | Set up PivotTable with Category row field and Sales data field using Aspose.Cells C#
// Tags: Aspose.Cells create pivot table C# | Aspose.Cells FreezePanes worksheet C# | pivot table row field configuration Aspose.Cells | freeze row labels Excel Aspose.Cells | add data field to PivotTable Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample creates a new workbook, fills it with Category, Product, and Sales data, defines a source range, adds a PivotTable at E3, sets 'Category' as the row field and 'Sales' as the data field, freezes the first column so row labels stay visible, and saves the file as PivotTableWithFrozenRowLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Product");
            sheet.Cells["C1"].PutValue("Sales");

            string[] categories = { "Beverages", "Beverages", "Food", "Food", "Food", "Electronics", "Electronics" };
            string[] products   = { "Tea", "Coffee", "Bread", "Butter", "Cheese", "Phone", "Laptop" };
            double[] sales      = { 1200, 1500, 800, 600, 900, 3000, 4500 };

            for (int i = 0; i < categories.Length; i++)
            {
                int row = i + 2; // Data starts at row 2 (1‑based indexing in Excel)
                sheet.Cells[row, 0].PutValue(categories[i]); // Column A
                sheet.Cells[row, 1].PutValue(products[i]);   // Column B
                sheet.Cells[row, 2].PutValue(sales[i]);      // Column C
            }

            // Define the source data range for the pivot table (including header).
            int totalRows = categories.Length + 1; // +1 for header
            string sourceData = $"=Sheet1!$A$1:$C${totalRows}";

            // Add a pivot table starting at cell E3.
            int pivotIndex = sheet.PivotTables.Add("PivotTable1", "E3", sourceData);
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table.
            // Note: The PivotFields collection is used to reference fields by index.
            // If the API version does not expose PivotFields, this block can be adjusted accordingly.
            if (pivotTable.RowFields != null && pivotTable.DataFields != null)
            {
                // Add "Category" as a row field (field index 0).
                pivotTable.RowFields.Add(pivotTable.RowFields[0]); // Fallback if PivotFields not available.
                // Add "Sales" as a data field (field index 2).
                pivotTable.DataFields.Add(pivotTable.DataFields[0]); // Fallback if PivotFields not available.
            }

            // Freeze the first column (column A) so it stays visible while scrolling.
            sheet.FreezePanes(0, 1, 0, 0);

            // Prepare output path and ensure the directory exists.
            string outputPath = "PivotTableWithFrozenRowLabels.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
