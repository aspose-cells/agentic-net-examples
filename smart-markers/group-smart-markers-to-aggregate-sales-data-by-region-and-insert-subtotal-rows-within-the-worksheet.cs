// Title: C# – Add Region‑Based Subtotal Rows to a Smart‑Marker Worksheet with Aspose.Cells
// Description: Creates a new workbook, defines headers, inserts smart‑marker placeholders for Region, Product, and Sales, binds a List<Sale> data source, processes the markers, calculates the populated range, and uses Cells.Subtotal to group rows by the Region column, sum the Sales values, insert page breaks, and place summary rows below each group before saving the file.
// Keywords: Aspose.Cells | C# smart markers | Excel subtotal | group by region | sales aggregation | Cells.Subtotal example | .NET Excel automation | regional totals | page break Excel | outline summary row | GitHub Aspose.Cells sample | US developers | European developers
// Common Searches: Aspose.Cells add subtotal rows after smart marker processing | C# group smart‑marker data by column and sum values | How to use Cells.Subtotal with smart markers in .NET | Create regional subtotals in Excel using Aspose.Cells | Smart marker template with automatic totals
// Developer Intent: Generate an Excel report where data inserted via smart markers is automatically grouped by region and each group shows a summed sales subtotal.
// Use Cases: Produce a sales report that groups entries by region and displays total sales per region. | Build a reusable Excel template that fills data through smart markers and adds regional subtotals for printing or distribution. | Create worksheets with page breaks between regions and summary rows for easier navigation and review.
// AI Prompts: Write C# code that uses Aspose.Cells to insert subtotal rows grouped by a specific column after smart markers are processed. | Explain how to change the Subtotal method parameters to place the summary row above the detail rows and disable page breaks. | Show how to capture the row indices created by the Subtotal operation for further custom formatting.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsSmartMarkerSubtotalDemo
{
    // Simple data model for sales records
    // Creates a new workbook, defines headers, inserts smart‑marker placeholders for Region, Product, and Sales, binds a List<Sale> data source, processes the markers, calculates the populated range, and uses Cells.Subtotal to group rows by the Region column, sum the Sales values, insert page breaks, and place summary rows below each group before saving the file.
    public class Sale
    {
        public string? Region { get; set; }
        public string? Product { get; set; }
        public double Sales { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // 2. Set up header row
                cells["A1"].PutValue("Region");
                cells["B1"].PutValue("Product");
                cells["C1"].PutValue("Sales");

                // 3. Insert smart markers for the data rows (starting at row 2)
                //    These markers will be replaced by the data source during processing
                cells["A2"].PutValue("&=$Region");
                cells["B2"].PutValue("&=$Product");
                cells["C2"].PutValue("&=$Sales");

                // 4. Define the range that contains the smart markers and give it the required name
                //    Aspose.Cells looks for a range named "_CellsSmartMarkers" when processing
                Aspose.Cells.Range smartMarkerRange = cells.CreateRange("A2:C2");
                smartMarkerRange.Name = "_CellsSmartMarkers";

                // 5. Prepare sample sales data
                List<Sale> salesData = new List<Sale>
                {
                    new Sale { Region = "North", Product = "Widget", Sales = 5000 },
                    new Sale { Region = "North", Product = "Gadget", Sales = 3000 },
                    new Sale { Region = "South", Product = "Widget", Sales = 6000 },
                    new Sale { Region = "South", Product = "Gadget", Sales = 4000 },
                    new Sale { Region = "West",  Product = "Widget", Sales = 4500 }
                };

                // 6. Create a WorkbookDesigner, assign the workbook and the data source, then process
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("Data", salesData);
                designer.Process(); // processes all smart markers in the workbook

                // 7. After processing, determine the total number of rows (header + data)
                int totalRows = salesData.Count + 1; // +1 for header row

                // 8. Define the cell area that includes the header and all data rows
                //    Columns A (0) to C (2), rows 0 to totalRows-1 (zero‑based)
                CellArea dataArea = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = totalRows - 1,
                    EndColumn = 2
                };

                // 9. Add subtotal rows:
                //    - Group by the first column (Region) -> groupBy = 0
                //    - Use SUM function on the Sales column (index 2)
                //    - Replace existing subtotals, add page breaks, place summary below data
                cells.Subtotal(
                    dataArea,
                    0,
                    ConsolidationFunction.Sum,
                    new int[] { 2 },
                    true,   // replace existing subtotals
                    true,   // add page breaks between groups
                    true    // place summary row below the detail rows
                );

                // 10. Ensure the outline shows the summary row below the grouped data
                worksheet.Outline.SummaryRowBelow = true;

                // 11. Save the resulting workbook
                workbook.Save("SmartMarkerSubtotalDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
