using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeBlockDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX template (ensure the file exists at the specified path)
            Workbook workbook = new Workbook("template.xlsx");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define a smart marker range block:
            // Row 0, Column 0 will contain the start token {{RANGE}}
            // Subsequent rows will contain data markers (e.g., &=$Name, &=$Value)
            // The last row will contain the end token {{ENDRANGE}}
            cells["A1"].PutValue("{{RANGE}}");
            cells["A2"].PutValue("&=$Name");   // Smart marker for Name
            cells["B2"].PutValue("&=$Value");  // Smart marker for Value
            cells["A3"].PutValue("{{ENDRANGE}}"); // End token

            // Create a Range object that covers the entire block (A1:B3)
            AsposeRange smartRange = cells.CreateRange("A1", "B3");
            // Assign a name recognized by WorkbookDesigner for smart markers
            smartRange.Name = "_CellsSmartMarkers";

            // Prepare a data source (list of anonymous objects)
            var data = new List<object>
            {
                new { Name = "Alice", Value = 123 },
                new { Name = "Bob",   Value = 456 },
                new { Name = "Carol", Value = 789 }
            };

            // Initialize WorkbookDesigner and bind the data source
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("Data", data);

            // Process only the defined range block (true = preserve unrecognized markers)
            designer.Process(smartRange, true);

            // Save the processed workbook
            workbook.Save("output.xlsx");
        }
    }
}