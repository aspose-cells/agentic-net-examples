// Title: Create a Pivot Table and Save with LightCellsDataProvider using Aspose.Cells for .NET
// Description: This C# example shows how to build a small data set, define a pivot table (Region → Rows, Product → Columns, Sales → Data), apply a tabular layout, disable pivot cache storage, and persist the workbook as an XLSX file using Aspose.Cells LightCellsDataProvider for memory‑efficient saving.
// Keywords: Aspose.Cells pivot table C# | LightCellsDataProvider save workbook | disable pivot cache Aspose | tabular layout pivot Aspose.Cells | memory efficient Excel export .NET | create pivot table programmatically | Aspose.Cells open workbook example
// Common Searches: how to add a pivot table with Aspose.Cells in C# | Aspose.Cells LightCellsDataProvider example | save Excel file without pivot cache using Aspose | C# code for pivot table rows columns data fields | Aspose.Cells create analytical view Excel
// Developer Intent: Generate a pivot table from worksheet data and save the workbook efficiently with LightCellsDataProvider.
// Use Cases: Produce a sales summary pivot that groups amounts by region and product. | Create an analytical view for quick reporting while keeping file size low. | Export large Excel workbooks with pivot tables using a memory‑optimized provider.
// AI Prompts: Generate C# code that builds a pivot table with Aspose.Cells, sets row/column/data fields, applies a tabular layout, disables cache storage, and saves the file using LightCellsDataProvider. | Explain step‑by‑step how to construct the source data range string for a pivot table in Aspose.Cells and export the workbook with LightCellsDataProvider. | Show how to configure Aspose.Cells PivotTable to hide cache data and reduce file size when saving with LightCellsDataProvider.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotWithLightCells
{
    // This C# example shows how to build a small data set, define a pivot table (Region → Rows, Product → Columns, Sales → Data), apply a tabular layout, disable pivot cache storage, and persist the workbook as an XLSX file using Aspose.Cells LightCellsDataProvider for memory‑efficient saving.
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Create a new workbook and add source data sheet
                // -------------------------------------------------
                Workbook workbook = new Workbook();                     // create workbook
                Worksheet sourceSheet = workbook.Worksheets[0];         // default first sheet
                sourceSheet.Name = "SourceData";

                // Populate sample data for the pivot table
                sourceSheet.Cells["A1"].PutValue("Region");
                sourceSheet.Cells["B1"].PutValue("Product");
                sourceSheet.Cells["C1"].PutValue("Sales");

                sourceSheet.Cells["A2"].PutValue("North");
                sourceSheet.Cells["B2"].PutValue("Apple");
                sourceSheet.Cells["C2"].PutValue(1200);

                sourceSheet.Cells["A3"].PutValue("North");
                sourceSheet.Cells["B3"].PutValue("Banana");
                sourceSheet.Cells["C3"].PutValue(850);

                sourceSheet.Cells["A4"].PutValue("South");
                sourceSheet.Cells["B4"].PutValue("Apple");
                sourceSheet.Cells["C4"].PutValue(950);

                sourceSheet.Cells["A5"].PutValue("South");
                sourceSheet.Cells["B5"].PutValue("Banana");
                sourceSheet.Cells["C5"].PutValue(1100);

                // -------------------------------------------------
                // 2. Add a worksheet that will host the pivot table
                // -------------------------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Build the source data reference string (e.g., =SourceData!$A$1:$C$5)
                string sourceData = $"=SourceData!{sourceSheet.Cells.MaxDisplayRange.Address}";

                // -------------------------------------------------
                // 3. Create the pivot table
                // -------------------------------------------------
                PivotTableCollection pivotTables = pivotSheet.PivotTables;
                int pivotIndex = pivotTables.Add(sourceData, "A1", "SalesPivot"); // dest cell A1, table name "SalesPivot"
                PivotTable pivotTable = pivotTables[pivotIndex];

                // Configure fields: Region -> Row, Product -> Column, Sales -> Data
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
                pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Optional: set layout to tabular form for clearer analytical view
                pivotTable.ShowInTabularForm();

                // Optional: prevent pivot data from being stored inside the file (saves space)
                pivotTable.SaveData = false;

                // -------------------------------------------------
                // 4. Save the workbook
                // -------------------------------------------------
                string outputPath = "PivotWithLightCells.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
