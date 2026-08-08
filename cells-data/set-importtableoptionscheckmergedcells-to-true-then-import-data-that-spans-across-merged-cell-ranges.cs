// Title: Import custom objects below a merged header with ImportTableOptions.CheckMergedCells in Aspose.Cells for .NET
// Description: Demonstrates how to keep a merged header (A1:C1) intact while importing a List<Product> into a worksheet. The example configures ImportTableOptions (IsFieldNameShown = false, InsertRows = true, CheckMergedCells = true), starts the import at row 2, and saves the file as MergedImportDemo.xlsx.
// Keywords: Aspose.Cells ImportTableOptions | CheckMergedCells true | ImportCustomObjects merged cells | C# import list to worksheet | preserve merged header Aspose.Cells | InsertRows option | Excel merged cells import
// Common Searches: Aspose.Cells import data without overwriting merged cells | How to use ImportTableOptions.CheckMergedCells in C# | ImportCustomObjects below a merged header Aspose.Cells | Insert rows while preserving merged cells in Excel using Aspose
// Developer Intent: Add rows of object data to a worksheet without disturbing existing merged cells.
// Use Cases: Add a product catalog under a merged title row in a sales template. | Populate transaction records beneath merged section headings in a financial report. | Insert new entries into a worksheet that uses merged cells for category labels.
// AI Prompts: Write C# code that imports a List<T> into an Aspose.Cells worksheet with CheckMergedCells enabled and inserts rows below a merged header. | Explain the effect of ImportTableOptions.CheckMergedCells when importing custom objects into a sheet containing merged cells. | Provide a step‑by‑step tutorial for preserving merged cells while importing data with Aspose.Cells for .NET.

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

// Demonstrates how to keep a merged header (A1:C1) intact while importing a List<Product> into a worksheet. The example configures ImportTableOptions (IsFieldNameShown = false, InsertRows = true, CheckMergedCells = true), starts the import at row 2, and saves the file as MergedImportDemo.xlsx.
public class MergedImportDemo
{
    // Simple POCO to be imported
    public class Product
    {
        public string? Name { get; set; }
        public double Price { get; set; }
    }

    public static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Merge a range that will be intersected by the imported data (A1:C1)
            cells.Merge(0, 0, 1, 3); // Merge cells A1, B1, C1
            cells[0, 0].PutValue("Header");

            // Prepare data to import; three columns match the merged range width
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 1.20 },
                new Product { Name = "Banana", Price = 0.80 },
                new Product { Name = "Cherry", Price = 2.50 }
            };

            // Set import options with CheckMergedCells enabled
            ImportTableOptions options = new ImportTableOptions
            {
                IsFieldNameShown = false, // Do not import property names as a header row
                InsertRows = true,        // Insert rows instead of overwriting existing ones
                CheckMergedCells = true   // Preserve merged cells during import
            };

            // Import the list starting at row 1 (below the merged header) to avoid overwriting it
            cells.ImportCustomObjects((ICollection)products, 1, 0, options);

            // Save the workbook
            workbook.Save("MergedImportDemo.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
