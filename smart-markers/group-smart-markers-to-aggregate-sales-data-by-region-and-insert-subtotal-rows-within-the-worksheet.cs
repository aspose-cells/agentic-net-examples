using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace SmartMarkerSubtotalDemo
{
    // Simple data model for sales records
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
                // Prepare sample sales data
                List<Sale> salesData = new List<Sale>
                {
                    new Sale { Region = "North", Product = "Widget", Sales = 5000 },
                    new Sale { Region = "North", Product = "Gadget", Sales = 3000 },
                    new Sale { Region = "South", Product = "Widget", Sales = 6000 },
                    new Sale { Region = "South", Product = "Gadget", Sales = 4000 },
                    new Sale { Region = "West",  Product = "Widget", Sales = 4500 }
                };

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Set up header row
                cells["A1"].PutValue("Region");
                cells["B1"].PutValue("Product");
                cells["C1"].PutValue("Sales");

                // Insert smart markers for the data rows (starting at row 2)
                cells["A2"].PutValue("&=$Region");
                cells["B2"].PutValue("&=$Product");
                cells["C2"].PutValue("&=$Sales");

                // Define the range that contains the smart markers and give it the required name
                AsposeRange smRange = cells.CreateRange("A2:C2");
                smRange.Name = "_CellsSmartMarkers";

                // Set up the WorkbookDesigner with the data source and process the smart markers
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("SalesData", salesData);
                designer.Process();

                // Determine the used range after processing
                int lastRow = cells.MaxDataRow;          // includes header row
                int lastColumn = cells.MaxDataColumn;    // should be 2 (C column)

                // Define the cell area that includes the header and all data rows
                CellArea dataArea = new CellArea
                {
                    StartRow = 0,          // header row (A1)
                    StartColumn = 0,       // column A
                    EndRow = lastRow,      // last populated row
                    EndColumn = lastColumn // last populated column
                };

                // Apply subtotal: group by Region (column 0), sum Sales (column 2)
                // Replace existing subtotals, add page breaks, place summary below data
                cells.Subtotal(
                    dataArea,
                    0,                                 // group by first column (Region)
                    ConsolidationFunction.Sum,         // use SUM function
                    new int[] { 2 },                   // subtotal on Sales column
                    true,                              // replace existing subtotals
                    true,                              // add page breaks between groups
                    true                               // place summary row below each group
                );

                // Ensure the summary rows appear below the detail rows
                worksheet.Outline.SummaryRowBelow = true;

                // Prepare output path and ensure directory exists
                string outputPath = "SmartMarkerSubtotalDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? "";
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the resulting workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}