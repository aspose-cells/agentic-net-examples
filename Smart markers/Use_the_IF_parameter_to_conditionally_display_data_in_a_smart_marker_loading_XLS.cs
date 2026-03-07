using System;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerIfDemo
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers with IF condition.
            // Example smart marker in the template (cell A2): &IF($Quantity>0,$Quantity,0)
            Workbook workbook = new Workbook("template.xlsx");

            // Initialize the WorkbookDesigner with the loaded workbook.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Prepare JSON data source. The JSON must contain the fields referenced in the IF expression.
            // In this example we have fields: Product, Quantity, Price.
            string jsonData = @"
            [
                { ""Product"": ""Apple"",  ""Quantity"": 10, ""Price"": 1.2 },
                { ""Product"": ""Banana"", ""Quantity"": 0,  ""Price"": 0.8 },
                { ""Product"": ""Cherry"", ""Quantity"": 5, ""Price"": 2.5 }
            ]";

            // Set the JSON data source with a name that matches the smart marker table name in the template.
            // Assuming the template uses the table name "Products" (e.g., &Products.Product, &IF($Quantity>0,$Quantity,0))
            designer.SetJsonDataSource("Products", jsonData);

            // Process the smart markers. The IF condition will evaluate for each row.
            designer.Process();

            // Save the result workbook.
            workbook.Save("output.xlsx");
        }
    }
}