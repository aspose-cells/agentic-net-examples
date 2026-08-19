// Title: Aspose.Cells WorkbookDesigner – Preserve Leading Apostrophes (QuotePrefixToStyle = false) in .NET
// Description: Demonstrates how to disable the QuotePrefixToStyle setting in Aspose.Cells so that leading apostrophes remain literal text when processing smart markers. The C# example creates a workbook, inserts a value with a leading apostrophe, adds a smart marker, binds a dictionary data source, runs WorkbookDesigner.Process(), and saves the result.
// Keywords: Aspose.Cells WorkbookDesigner | QuotePrefixToStyle false | preserve leading apostrophe | smart markers apostrophe handling | .NET Excel export | keep apostrophe in cell value | Excel template literal apostrophe
// Common Searches: Aspose.Cells keep leading apostrophe | WorkbookDesigner disable QuotePrefixToStyle | prevent apostrophe removal in Excel export | smart marker literal apostrophe C# | how to retain apostrophe in Aspose.Cells
// Developer Intent: Configure WorkbookDesigner to treat a leading apostrophe as a regular character rather than a formatting prefix.
// Use Cases: Export product codes or identifiers that start with an apostrophe (e.g., '001) without losing the character. | Generate Excel reports from templates that contain static text or smart markers with leading apostrophes. | Preserve user‑entered strings containing apostrophes when creating or updating Excel files programmatically.
// AI Prompts: Show how to set QuotePrefixToStyle = false in Aspose.Cells WorkbookDesigner to keep leading apostrophes. | Provide a C# snippet that inserts a cell value with a leading apostrophe, processes smart markers, and verifies the apostrophe remains. | Explain the impact of the QuotePrefixToStyle property on cell formatting and when to toggle it for smart‑marker operations.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsWorkbookDesignerQuotePrefixDemo
{
    // Demonstrates how to disable the QuotePrefixToStyle setting in Aspose.Cells so that leading apostrophes remain literal text when processing smart markers. The C# example creates a workbook, inserts a value with a leading apostrophe, adds a smart marker, binds a dictionary data source, runs WorkbookDesigner.Process(), and saves the result.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing template)
            Workbook workbook = new Workbook();

            // IMPORTANT: Disable QuotePrefixToStyle so leading apostrophes are kept as literal characters
            workbook.Settings.QuotePrefixToStyle = false;

            // Add some sample data that includes a leading apostrophe
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("'SampleText"); // apostrophe should be part of the value

            // Insert a simple smart marker for demonstration
            sheet.Cells["B1"].PutValue("&=Data.Value");

            // Prepare a data source that will be bound to the smart marker
            var data = new Dictionary<string, object>
            {
                { "Value", "BoundValue" }
            };

            // Initialize WorkbookDesigner with the workbook
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Bind the data source (name "Data" matches the smart marker prefix)
            designer.SetDataSource("Data", data);

            // Process the smart markers
            designer.Process();

            // Save the resulting workbook
            workbook.Save("WorkbookDesigner_WithLiteralApostrophe.xlsx");
        }
    }
}
