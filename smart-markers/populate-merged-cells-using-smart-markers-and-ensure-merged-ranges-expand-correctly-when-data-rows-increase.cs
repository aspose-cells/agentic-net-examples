// Title: How to populate merged header cells and auto‑expand smart‑marker rows in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates an Excel workbook, merges a header row across columns, inserts a smart‑marker row, binds a List<Product> to WorkbookDesigner, and processes the markers to duplicate rows for each product. | Show how to keep the merged header unchanged while the smart‑marker range expands dynamically based on the data source size in Aspose.Cells. | Demonstrate applying custom styling to a merged header after processing smart markers and saving the workbook as an .xlsx file.
// Common Searches: aspnet c# using Aspose.Cells smart markers with merged header rows | expand smart marker rows automatically when binding a list in Aspose.Cells | preserve merged cells while processing smart markers in Excel using Aspose.Cells | naming convention for smart marker ranges in Aspose.Cells designer | example of WorkbookDesigner processing a List<T> with merged header in C#
// Tags: merge header cells Aspose.Cells | smart marker row expansion WorkbookDesigner | CellsSmartMarkers range naming convention | populate Excel from C# List Aspose.Cells | maintain merged header during smart marker processing

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsSmartMarkersMergedDemo
{
    // Sample data class
    // C# example that creates a workbook with a merged header, defines a smart‑marker range, binds a List<Product> to WorkbookDesigner, processes the markers to auto‑expand data rows while keeping the merged header intact, applies header styling, and saves the file as SmartMarkersMergedDemo.xlsx.
    public class Product
    {
        public string Category { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public double Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a new workbook (lifecycle rule: create) ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ---------- Define a merged header that should span the data columns ----------
                // Merge cells A1:C1 to create a header area
                cells.Merge(0, 0, 1, 3);
                cells[0, 0].PutValue("Product List");
                // Apply a simple style to the merged header
                Style headerStyle = cells[0, 0].GetStyle();
                headerStyle.HorizontalAlignment = TextAlignmentType.Center;
                headerStyle.VerticalAlignment = TextAlignmentType.Center;
                headerStyle.Font.IsBold = true;
                cells[0, 0].SetStyle(headerStyle);

                // ---------- Add column titles ----------
                cells[1, 0].PutValue("Category");
                cells[1, 1].PutValue("Name");
                cells[1, 2].PutValue("Price");

                // ---------- Insert smart markers for the data rows ----------
                // Smart markers are placed in the row that will be repeated for each data item.
                // The range A3:C3 will be duplicated for each Product in the data source.
                cells[2, 0].PutValue("&=Category");
                cells[2, 1].PutValue("&=Name");
                cells[2, 2].PutValue("&=Price");

                // Name the range that contains the smart markers.
                // The name must start with an underscore and end with "_CellsSmartMarkers"
                // so that WorkbookDesigner recognises it automatically.
                AsposeRange dataRange = cells.CreateRange(2, 0, 1, 3);
                dataRange.Name = "_DataRange_CellsSmartMarkers";

                // ---------- Prepare sample data ----------
                List<Product> products = new List<Product>
                {
                    new Product { Category = "Beverages", Name = "Tea",      Price = 1.20 },
                    new Product { Category = "Beverages", Name = "Coffee",   Price = 2.50 },
                    new Product { Category = "Snacks",    Name = "Cookies",  Price = 3.10 },
                    new Product { Category = "Snacks",    Name = "Chips",    Price = 1.80 }
                };

                // ---------- Set up WorkbookDesigner and bind the data source ----------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                // Bind the list of products to the designer.
                designer.SetDataSource("Products", products);

                // ---------- Process smart markers (lifecycle rule: process) ----------
                // The designer will expand the range containing the smart markers,
                // inserting as many rows as there are items in the data source.
                // Merged cells defined outside the smart‑marker range (the header) remain intact.
                designer.Process();

                // ---------- Save the result (lifecycle rule: save) ----------
                string outputPath = "SmartMarkersMergedDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
