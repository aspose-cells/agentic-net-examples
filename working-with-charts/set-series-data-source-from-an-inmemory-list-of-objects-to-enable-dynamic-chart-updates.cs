// Title: Create and refresh an Aspose.Cells chart from an in‑memory List<T> in C#
// Common Searches: how to populate an Aspose.Cells worksheet from a List of objects in C# | bind Aspose.Cells chart series to a range generated from a List<T> | update Aspose.Cells chart data after modifying the underlying collection | refresh chart series values programmatically with Aspose.Cells .NET | dynamic Excel chart using Aspose.Cells and a C# object list
// Tags: Aspose.Cells populate worksheet from List<T> | Aspose.Cells bind chart series to cell range | Aspose.Cells update NSeries values dynamically | C# dynamic chart data source Aspose.Cells | Excel chart refresh after collection change Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example writes a List<Product> to a worksheet, creates a column chart bound to the written range, saves the workbook, then modifies the list, rewrites the data, updates the chart's NSeries ranges, and saves the updated workbook, demonstrating dynamic chart updates with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // ------------------------------------------------------------
            // 1. Prepare an in‑memory list of objects (the data source)
            // ------------------------------------------------------------
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Sales = 120 },
                new Product { Name = "Banana", Sales =  80 },
                new Product { Name = "Cherry", Sales = 150 }
            };

            // ------------------------------------------------------------
            // 2. Create a workbook and write the list into a worksheet
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Sales");

            // Write the list data starting from row 2 (index 1)
            WriteProductsToSheet(cells, products, startRow: 1);

            // ------------------------------------------------------------
            // 3. Add a chart and bind it to the data range written above
            // ------------------------------------------------------------
            // Add a column chart positioned at rows 5‑20 and columns 0‑8
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIdx];

            // Define the data ranges for values and categories
            int dataStartRow = 2; // Excel rows are 1‑based; row 2 = first data row
            int dataEndRow   = dataStartRow + products.Count - 1;

            string salesRange    = $"B{dataStartRow}:B{dataEndRow}";
            string categoryRange = $"A{dataStartRow}:A{dataEndRow}";

            // Bind the series to the sales column (values) and set category labels
            chart.NSeries.Add(salesRange, true);
            chart.NSeries.CategoryData = categoryRange;

            // Optional: give the series a friendly name
            chart.NSeries[0].Name = "Sales";

            // Save the workbook with the initial chart
            workbook.Save("ChartFromList_initial.xlsx");
            Console.WriteLine("Initial chart created → ChartFromList_initial.xlsx");

            // ------------------------------------------------------------
            // 4. Simulate a dynamic update: modify the in‑memory list
            // ------------------------------------------------------------
            products.Add(new Product { Name = "Date", Sales = 200 }); // new item
            products[1].Sales = 95;                                 // update Banana

            // Clear previous data rows (keep the header)
            int maxRow = sheet.Cells.MaxDataRow;
            for (int r = 1; r <= maxRow; r++)
            {
                sheet.Cells[r, 0].PutValue(string.Empty);
                sheet.Cells[r, 1].PutValue(string.Empty);
            }

            // Write the updated list back to the sheet
            WriteProductsToSheet(cells, products, startRow: 1);

            // Re‑calculate the data range after the list change
            dataEndRow   = dataStartRow + products.Count - 1;
            salesRange    = $"B{dataStartRow}:B{dataEndRow}";
            categoryRange = $"A{dataStartRow}:A{dataEndRow}";

            // Update the chart's series to point to the new ranges
            chart.NSeries[0].Values = salesRange;
            chart.NSeries.CategoryData = categoryRange;

            // Save the workbook with the updated chart
            workbook.Save("ChartFromList_updated.xlsx");
            Console.WriteLine("Chart updated with new data → ChartFromList_updated.xlsx");
        }
        catch (Exception ex)
        {
            // Any unexpected error will be printed to the console.
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    /// <param name="cells">The Cells collection of the target worksheet.</param>
    /// <param name="products">The data source.</param>
    /// <param name="startRow">Zero‑based row index where writing should begin (excluding header).</param>
    static void WriteProductsToSheet(Cells cells, List<Product> products, int startRow)
    {
        for (int i = 0; i < products.Count; i++)
        {
            cells[startRow + i, 0].PutValue(products[i].Name);
            cells[startRow + i, 1].PutValue(products[i].Sales);
        }
    }
}

/// <summary>
/// Simple POCO representing a product and its sales figure.
/// </summary>
class Product
{
    public string Name { get; set; }
    public double Sales { get; set; }
}
