using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace ConditionalFormattingSmartMarkerDemo
{
    // Simple data class used as a data source for smart markers
    public class Product
    {
        public string Name { get; set; }
        public double Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Add header cells with smart markers that will be populated from the data source
            //    The smart marker syntax "&=Products.Name" and "&=Products.Price" will be replaced.
            sheet.Cells["A1"].PutValue("Product Name");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("&=Products.Name");
            sheet.Cells["B2"].PutValue("&=Products.Price");

            // 3. Define a conditional formatting rule that highlights prices greater than 50
            //    The rule will be applied to the range B2:B10 (where prices will be inserted).
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

            // Define the area for the conditional formatting
            CellArea priceArea = new CellArea
            {
                StartRow = 1,   // Row 2 (zero‑based)
                EndRow = 9,     // Row 10
                StartColumn = 1, // Column B
                EndColumn = 1
            };
            fcs.AddArea(priceArea);

            // Add a condition: CellValue > 50
            int conditionIdx = fcs.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "50", null);
            FormatCondition priceCondition = fcs[conditionIdx];
            priceCondition.Style.BackgroundColor = System.Drawing.Color.LightSalmon;
            priceCondition.Style.Font.Color = System.Drawing.Color.Black;

            // 4. Prepare sample data source
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple", Price = 30 },
                new Product { Name = "Banana", Price = 55 },
                new Product { Name = "Cherry", Price = 70 },
                new Product { Name = "Date", Price = 45 }
            };

            // 5. Set up WorkbookDesigner, bind the data source, and process smart markers
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("Products", products);
            designer.Process(); // processes all smart markers in the workbook

            // 6. (Optional) Verify that smart markers have been replaced and conditional formatting is intact
            Console.WriteLine("Smart markers after processing:");
            foreach (string marker in designer.GetSmartMarkers())
            {
                Console.WriteLine(marker);
            }

            // 7. Save the resulting workbook
            workbook.Save("ConditionalFormattingWithSmartMarkers.xlsx");
        }
    }
}