using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerSmartMarkerDemo
{
    public class Sale
    {
        public string Category { get; set; }
        public string Product { get; set; }
        public double Amount { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            string templatePath = "TemplateWithSlicerSmartMarkers.xlsx";

            // If the template does not exist, create a minimal one with a smart marker.
            if (!File.Exists(templatePath))
            {
                Workbook tempWb = new Workbook();
                Worksheet ws = tempWb.Worksheets[0];
                ws.Name = "Data";

                // Place a smart marker that uses slicer filtering.
                ws.Cells["A1"].PutValue("&=[Sales].Slicer(Category).Amount");

                // Save the template.
                tempWb.Save(templatePath);
            }

            // Load the template workbook that contains smart markers.
            Workbook workbook = new Workbook(templatePath);

            // Create a WorkbookDesigner and attach the loaded workbook.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Prepare a list of sales data.
            List<Sale> sales = new List<Sale>
            {
                new Sale { Category = "Fruit",    Product = "Apple",   Amount = 120.5 },
                new Sale { Category = "Fruit",    Product = "Banana",  Amount = 85.0 },
                new Sale { Category = "Beverage",Product = "Coffee",  Amount = 150.0 },
                new Sale { Category = "Beverage",Product = "Tea",     Amount = 95.5 },
                new Sale { Category = "Snack",   Product = "Chips",   Amount = 60.0 }
            };

            // Set the data source for the smart marker.
            designer.SetDataSource("Sales", sales);

            // Process the smart markers.
            designer.Process();

            // Save the result workbook.
            workbook.Save("ResultWithFilteredData.xlsx");
        }
    }
}