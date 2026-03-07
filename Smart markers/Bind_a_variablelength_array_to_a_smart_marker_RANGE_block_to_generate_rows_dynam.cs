using System;
using System.Collections.Generic;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsSmartMarkerDemo
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook (XLSX) that contains a smart marker range.
            Workbook workbook = new Workbook("template.xlsx");

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define the range that contains the smart markers.
            AsposeRange smartMarkerRange = cells.CreateRange("A2:K2");
            smartMarkerRange.Name = "_CellsSmartMarkers";

            // Prepare a variable‑length data source.
            var items = new List<object>
            {
                new { Name = "Alice", Age = 30, Country = "USA" },
                new { Name = "Bob", Age = 25, Country = "Canada" },
                new { Name = "Charlie", Age = 35, Country = "UK" }
            };

            // Create a WorkbookDesigner and bind the data source.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("Items", items);

            // Process only the defined smart marker range.
            designer.Process(smartMarkerRange, true);

            // Save the result workbook.
            workbook.Save("output.xlsx");
        }
    }
}