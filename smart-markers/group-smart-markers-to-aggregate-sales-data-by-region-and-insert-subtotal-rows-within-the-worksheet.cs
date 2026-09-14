// Title: Generate an Excel file with Aspose.Cells C# smart markers, group rows by Region and add subtotal rows for Sales
// AI Prompts: Write C# code that creates a Workbook, inserts smart markers for a List<SaleRecord>, processes them with WorkbookDesigner, then calls Cells.Subtotal to group by the Region column and sum the Sales column, placing the subtotal rows below each group. | Show how to configure the worksheet outline so that the subtotal rows appear below the detail rows after applying the Cells.Subtotal method in Aspose.Cells.
// Common Searches: aspnet core aspose.cells smart markers group by column and add subtotal rows | c# how to insert subtotal rows after processing smart markers with Aspose.Cells | aspose.cells subtotal function for grouped sales data in Excel using C#
// Tags: Aspose.Cells smart markers binding | group by column subtotal Aspose.Cells | insert subtotal rows Excel C# | worksheet outline summary row Aspose.Cells | sales data aggregation Excel Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace SmartMarkerSubtotalDemo
{
    // Simple POCO representing a sales record
    // Demonstrates creating a workbook, defining smart markers for a list of SaleRecord objects, processing them with WorkbookDesigner, applying Cells.Subtotal to group by the Region column and sum the Sales column, configuring the outline to show summary rows below each group, and saving the result as an Excel file.
    public class SaleRecord
    {
        public string Region { get; set; } = string.Empty;
        public string Product { get; set; } = string.Empty;
        public double Sales { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Prepare sample sales data
                List<SaleRecord> sales = new List<SaleRecord>
                {
                    new SaleRecord { Region = "North", Product = "Widget", Sales = 5000 },
                    new SaleRecord { Region = "North", Product = "Gadget", Sales = 3000 },
                    new SaleRecord { Region = "South", Product = "Widget", Sales = 6000 },
                    new SaleRecord { Region = "South", Product = "Gadget", Sales = 4000 },
                    new SaleRecord { Region = "West",  Product = "Widget", Sales = 4500 }
                };

                // 2. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 3. Write header row
                cells["A1"].PutValue("Region");
                cells["B1"].PutValue("Product");
                cells["C1"].PutValue("Sales");

                // 4. Insert smart markers for the data rows (starting at row 2)
                cells["A2"].PutValue("&=$Region");
                cells["B2"].PutValue("&=$Product");
                cells["C2"].PutValue("&=$Sales");

                // 5. Define the range that contains the smart markers
                // The range must be named "_CellsSmartMarkers" for the designer to process it
                AsposeRange smRange = cells.CreateRange("A2:C2");
                smRange.Name = "_CellsSmartMarkers";

                // 6. Set up the WorkbookDesigner, bind the data source and process the smart markers
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("SalesData", sales);
                designer.Process(); // populates the rows based on the smart markers

                // 7. Determine the last row after processing (data rows start at row 2)
                int lastDataRow = cells.MaxDataRow; // includes the header row

                // 8. Define the cell area that covers the whole table (including header)
                CellArea area = CellArea.CreateCellArea(0, 0, lastDataRow, 2); // rows 0..lastDataRow, columns A..C

                // 9. Add subtotal rows: group by the first column (Region), sum the Sales column (index 2)
                //    replace = true, pageBreaks = false, summaryBelowData = true
                cells.Subtotal(
                    area,
                    0, // Group by column 0 (Region)
                    ConsolidationFunction.Sum,
                    new int[] { 2 }, // Apply subtotal to column 2 (Sales)
                    true,   // replace existing subtotals if any
                    false,  // do not insert page breaks between groups
                    true    // place summary rows below each group
                );

                // 10. Ensure the outline shows summary rows below the detail rows
                sheet.Outline.SummaryRowBelow = true;

                // 11. Save the resulting workbook
                string outputPath = "SmartMarkerSubtotalDemo.xlsx";
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
