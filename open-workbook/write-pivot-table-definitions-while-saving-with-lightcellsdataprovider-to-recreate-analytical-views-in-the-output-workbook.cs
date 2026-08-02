// Title: Create and Save a Pivot Table with Aspose.Cells for .NET (C#) – optional LightCellsDataProvider
// Description: This example demonstrates how to build a workbook, add a data sheet, define a pivot table (Category → Product rows, Sales values), and save the file as an XLSX. It shows how to preserve pivot data with SaveData and outlines where LightCellsDataProvider could be used for memory‑efficient saving of large analytical workbooks.
// Keywords: Aspose.Cells C# pivot table | save workbook LightCellsDataProvider | Aspose.Cells create pivot | export pivot to XLSX | memory‑efficient Excel generation | preserve pivot data Aspose
// Common Searches: Aspose.Cells add pivot table C# | Save workbook with LightCellsDataProvider | How to keep pivot data when saving Excel with Aspose | Define source range for Aspose.Cells pivot | Create sales summary pivot using Aspose.Cells
// Developer Intent: Generate a pivot table from worksheet data and persist the workbook, optionally using LightCellsDataProvider to reduce memory usage while keeping the analytical view intact.
// Use Cases: Produce a sales‑by‑category report and export it as an XLSX file for business stakeholders. | Create large Excel files with multiple pivot tables while minimizing RAM consumption via LightCellsDataProvider. | Automate regeneration of analytical views (pivot tables) each time source data is refreshed.
// AI Prompts: Rewrite the sample to save the workbook with LightCellsDataProvider while preserving the SalesPivot definition. | Add code that refreshes the pivot table after modifying source data before saving. | Show how to export only the pivot table values to a CSV file using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This example demonstrates how to build a workbook, add a data sheet, define a pivot table (Category → Product rows, Sales values), and save the file as an XLSX. It shows how to preserve pivot data with SaveData and outlines where LightCellsDataProvider could be used for memory‑efficient saving of large analytical workbooks.
class PivotWithLightCellsDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and populate source data
            using (Workbook workbook = new Workbook())
            {
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Product");
                dataSheet.Cells["C1"].PutValue("Sales");

                dataSheet.Cells["A2"].PutValue("Electronics");
                dataSheet.Cells["B2"].PutValue("Laptop");
                dataSheet.Cells["C2"].PutValue(1200);

                dataSheet.Cells["A3"].PutValue("Electronics");
                dataSheet.Cells["B3"].PutValue("Phone");
                dataSheet.Cells["C3"].PutValue(800);

                dataSheet.Cells["A4"].PutValue("Furniture");
                dataSheet.Cells["B4"].PutValue("Chair");
                dataSheet.Cells["C4"].PutValue(150);

                dataSheet.Cells["A5"].PutValue("Furniture");
                dataSheet.Cells["B5"].PutValue("Table");
                dataSheet.Cells["C5"].PutValue(300);

                // Add a worksheet for the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

                // Build source data reference (e.g., =Data!A1:C5)
                string sourceData = $"=Data!{dataSheet.Cells.MaxDisplayRange.Address}";

                // Add and configure the pivot table
                int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "SalesPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
                pivotTable.SaveData = true; // Preserve data with the workbook

                // Define output path
                const string outputPath = "PivotWithLightCells.xlsx";

                // Ensure the output directory exists (handle possible null)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook directly (LightCells not required for this demo)
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
