using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsNestedSmartMarkers
{
    // Sample data classes representing nested objects
    public class RootData
    {
        public string Title { get; set; }
        public List<ChildItem> Items { get; set; }
    }

    public class ChildItem
    {
        public string Name { get; set; }
        public double Price { get; set; }
        public DateTime ReleaseDate { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the Excel template that contains smart markers (e.g., &RootData.Title, &RootData.Items.Name, etc.)
            Workbook workbook = new Workbook("template.xlsx");

            // Prepare nested data source
            var data = new List<RootData>
            {
                new RootData
                {
                    Title = "Product Catalog",
                    Items = new List<ChildItem>
                    {
                        new ChildItem { Name = "Laptop", Price = 999.99, ReleaseDate = new DateTime(2023, 5, 1) },
                        new ChildItem { Name = "Smartphone", Price = 699.49, ReleaseDate = new DateTime(2023, 6, 15) },
                        new ChildItem { Name = "Tablet", Price = 399.00, ReleaseDate = new DateTime(2023, 7, 20) }
                    }
                }
            };

            // Initialize WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Bind the nested data source to the smart marker name used in the template
            designer.SetDataSource("RootData", data);

            // Process smart markers to populate the worksheet with nested object data
            designer.Process();

            // Save the populated workbook
            workbook.Save("output.xlsx");
        }
    }
}