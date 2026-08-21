// Title: Map Nested JSON to Excel with Smart Markers – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a Workbook, insert smart markers that reference nested JSON properties using dot notation, define a named range for a repeating collection, bind the JSON string via SetJsonDataSource, process the markers, and save the populated file as an XLSX document.
// Keywords: Aspose.Cells | C# | .NET | smart markers | nested JSON | JSON to Excel | SetJsonDataSource | WorkbookDesigner | repeatable range | Excel automation
// Common Searches: Aspose.Cells smart markers nested JSON example | C# map JSON object to Excel cells | populate Excel table from JSON array using Aspose | dot notation smart markers Aspose.Cells | repeat rows for JSON collection in Excel C#
// Developer Intent: Bind a hierarchical JSON document to an Excel workbook and fill cells with smart markers, including a dynamic table for array items.
// Use Cases: Insert employee name, street, and city into cells A1‑A3 using dot‑notation markers. | Generate a project list where each row repeats for every entry in Employee.Projects, showing title and budget. | Save the final workbook as JsonSmartMarkerOutput.xlsx after processing.
// AI Prompts: Write C# code that binds a nested JSON string to Aspose.Cells WorkbookDesigner and applies smart markers for both single values and a repeating collection. | Explain how to create a named smart‑marker range that expands for each element in a JSON array. | Provide debugging steps when smart markers return empty cells for nested JSON properties.

using System;
using Aspose.Cells;

namespace AsposeCellsJsonSmartMarkerDemo
{
    // Demonstrates how to create a Workbook, insert smart markers that reference nested JSON properties using dot notation, define a named range for a repeating collection, bind the JSON string via SetJsonDataSource, process the markers, and save the populated file as an XLSX document.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // 2. Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // 3. Insert smart markers that reference nested JSON properties
            //    The "$" prefix indicates a smart marker; nested properties are accessed with dot notation.
            sheet.Cells["A1"].PutValue("&=$DataSource.Employee.Name");
            sheet.Cells["A2"].PutValue("&=$DataSource.Employee.Address.Street");
            sheet.Cells["A3"].PutValue("&=$DataSource.Employee.Address.City");
            // Example of a collection (Projects). The marker will repeat for each item in the array.
            // Define a range that will be processed as a table of smart markers.
            sheet.Cells["A5"].PutValue("Project Title");
            sheet.Cells["B5"].PutValue("Budget");
            // Mark the start of the repeating range
            sheet.Cells.CreateRange("A6:B6").Name = "_CellsSmartMarkers";
            sheet.Cells["A6"].PutValue("&=$DataSource.Employee.Projects.Title");
            sheet.Cells["B6"].PutValue("&=$DataSource.Employee.Projects.Budget");

            // 4. Prepare JSON data with nested objects and an array
            string json = @"
            {
                ""Employee"": {
                    ""Name"": ""John Doe"",
                    ""Address"": {
                        ""Street"": ""123 Main St"",
                        ""City"": ""New York""
                    },
                    ""Projects"": [
                        { ""Title"": ""Project Alpha"", ""Budget"": 1500 },
                        { ""Title"": ""Project Beta"",  ""Budget"": 3000 }
                    ]
                }
            }";

            // 5. Initialize WorkbookDesigner and bind the JSON data source (rule: SetJsonDataSource)
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            // The first parameter is the name used in smart markers (DataSource in this case)
            designer.SetJsonDataSource("DataSource", json);

            // 6. Process the smart markers to populate the cells (rule: Process)
            designer.Process();

            // 7. Save the populated workbook (lifecycle: save)
            workbook.Save("JsonSmartMarkerOutput.xlsx");
        }
    }
}
