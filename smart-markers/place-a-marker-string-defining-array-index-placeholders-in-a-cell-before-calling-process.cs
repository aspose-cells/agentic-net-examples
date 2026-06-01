using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerArrayDemo
{
    // Simple data class representing a product
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Place a smart marker with an array index placeholder in a cell.
            //    The placeholder ${i} will be replaced by the current index during processing.
            //    Syntax: &=[DataSourceName][${index}].PropertyName
            //    Here we use "Products" as the data source name and "Name" as the property.
            sheet.Cells["A1"].PutValue("&=Products[${i}].Name");

            // 3. Prepare a list of products that will be used as the data source.
            List<Product> productList = new List<Product>
            {
                new Product { Name = "Apple",  Price = 1.20 },
                new Product { Name = "Banana", Price = 0.80 },
                new Product { Name = "Cherry", Price = 2.50 }
            };

            // 4. Create a WorkbookDesigner, assign the workbook and set the data source.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("Products", productList);

            // 5. Process the smart markers. The placeholder ${i} will be replaced with 0,1,2...
            designer.Process();

            // 6. Save the resulting workbook.
            workbook.Save("SmartMarkerArrayResult.xlsx");
        }
    }
}