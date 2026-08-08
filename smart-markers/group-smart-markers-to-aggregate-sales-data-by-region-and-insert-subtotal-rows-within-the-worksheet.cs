// Title: C# Aspose.Cells Smart Markers – Group Sales by Region and Add Subtotal Rows
// Description: Creates a workbook template with smart markers, fills it from a List<Sale> using WorkbookDesigner, then uses Cells.Subtotal to group rows by the Region column, sum the Sales column, and place summary rows below each group before saving.
// Keywords: Aspose.Cells | Smart Markers | C# | .NET | Subtotal rows | group by region | WorkbookDesigner | CellArea | outline summary row | Excel export
// Common Searches: Aspose.Cells smart markers subtotal example C# | group rows by column and add subtotals with Aspose.Cells | C# generate sales report with region totals using Aspose.Cells | how to use Cells.Subtotal in Aspose.Cells .NET | add outline summary rows below data Aspose.Cells
// Developer Intent: Populate an Excel sheet from a collection via smart markers and automatically insert region‑wise subtotal rows.
// Use Cases: Generate a sales report from a List<Sale> where each region’s sales are summed with a subtotal row. | Build a reusable Excel template that applies grouping and subtotals without manual formulas. | Export data with outline settings so summary rows appear directly beneath their detail rows.
// AI Prompts: Write C# code that uses Aspose.Cells WorkbookDesigner to replace smart markers with a collection and then adds subtotal rows grouped by a specified column. | Explain the required CellArea and Subtotal method parameters to replace existing subtotals and place summary rows below the detail rows. | Show how to configure the worksheet outline so that summary rows are displayed below the grouped data after adding subtotals.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerSubtotalDemo
{
    // Simple POCO representing a sales record
    // Creates a workbook template with smart markers, fills it from a List<Sale> using WorkbookDesigner, then uses Cells.Subtotal to group rows by the Region column, sum the Sales column, and place summary rows below each group before saving.
    public class Sale
    {
        public string Region { get; set; }
        public string Product { get; set; }
        public double Sales { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Set up the template with headers and smart markers
                cells["A1"].PutValue("Region");
                cells["B1"].PutValue("Product");
                cells["C1"].PutValue("Sales");

                // Smart markers – they will be replaced by the data source
                cells["A2"].PutValue("&=$Region");
                cells["B2"].PutValue("&=$Product");
                cells["C2"].PutValue("&=$Sales");

                // 3. Prepare sample sales data
                List<Sale> salesData = new List<Sale>
                {
                    new Sale { Region = "North", Product = "Widget", Sales = 5000 },
                    new Sale { Region = "North", Product = "Gadget", Sales = 3000 },
                    new Sale { Region = "South", Product = "Widget", Sales = 6000 },
                    new Sale { Region = "South", Product = "Gadget", Sales = 4000 },
                    new Sale { Region = "West",  Product = "Widget", Sales = 4500 }
                };

                // 4. Process the smart markers using WorkbookDesigner
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("SalesData", salesData);
                designer.Process(); // fills the template with the list data

                // 5. Determine the data range that now contains the populated rows
                int startRow = 0;               // header row (zero‑based)
                int startColumn = 0;            // column A
                int endRow = cells.MaxDataRow;  // last row with data after processing
                int endColumn = 2;              // column C (Sales)

                // 6. Create a CellArea covering the whole table (including header)
                CellArea dataArea = CellArea.CreateCellArea(startRow, startColumn, endRow, endColumn);

                // 7. Add subtotal rows: group by Region (column 0), sum Sales (column 2)
                //    Replace existing subtotals = true, no page breaks, summary placed below data = true
                cells.Subtotal(
                    dataArea,
                    0,                                 // group by first column (Region)
                    ConsolidationFunction.Sum,         // use SUM for subtotals
                    new int[] { 2 },                   // apply subtotal to Sales column
                    true,                              // replace existing subtotals
                    false,                             // do not insert page breaks between groups
                    true                               // place summary rows below the detail rows
                );

                // 8. Ensure the outline shows summary rows below the detail rows
                sheet.Outline.SummaryRowBelow = true;

                // 9. Save the resulting workbook
                string outputPath = "SmartMarkerSubtotalResult.xlsx";
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
