// Title: Insert a scalar smart marker into an Excel cell using Aspose.Cells for .NET
// Description: Demonstrates how to place the variable marker "&=Name" in cell A1, bind a C# Product object as the data source, process the marker with WorkbookDesigner, and save the workbook so the marker is replaced by the Product.Name value.
// Keywords: Aspose.Cells | smart marker | scalar value | C# | WorkbookDesigner | SetDataSource | Excel automation | variable marker | =&Name | dotnet example
// Common Searches: Aspose.Cells scalar smart marker example | C# place variable marker in Excel cell | WorkbookDesigner SetDataSource object | replace &=Name with property value | populate Excel cell from C# DTO
// Developer Intent: Bind a single property from a C# object to a smart‑marker cell in an Excel worksheet.
// Use Cases: Create a product label where the name is filled automatically from a data model. | Generate report headers that pull configuration values from a .NET DTO. | Build invoice templates that insert a single customer or company name via a smart marker.
// AI Prompts: Show how to add a scalar smart marker to a specific Excel cell and bind it to a C# property with Aspose.Cells. | Provide a step‑by‑step guide for using WorkbookDesigner to set a data source object and process a variable marker. | Explain how to handle multiple scalar smart markers in one worksheet using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using System.Collections.Generic;

namespace AsposeCellsSmartMarkerDemo
{
    // Simple data class representing a scalar value source
    // Demonstrates how to place the variable marker "&=Name" in cell A1, bind a C# Product object as the data source, process the marker with WorkbookDesigner, and save the workbook so the marker is replaced by the Product.Name value.
    public class Product
    {
        public string Name { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Place a smart marker variable in a cell (A1)
            // The marker "&=Name" will be replaced by the value of the "Name" property
            cells["A1"].PutValue("&=Name");

            // 3. Prepare the data source (scalar value)
            Product product = new Product { Name = "Aspose.Cells" };

            // 4. Create a WorkbookDesigner, assign the workbook and set the data source
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Product", product); // "Product" is the data source name

            // 5. Process the smart markers – this will replace the marker with the actual value
            designer.Process();

            // 6. Save the resulting workbook
            workbook.Save("SmartMarkerResult.xlsx");
        }
    }
}
