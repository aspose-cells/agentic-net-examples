// Title: List WebQuery Connections in Multiple Excel Workbooks with Aspose.Cells for .NET
// Description: The example loads each workbook from a supplied list, reads its DataConnections collection, filters for WebQueryConnection objects, and writes the connection name, URL and IsHtmlTables flag to the console, reporting when no web queries are present.
// Keywords: Aspose.Cells | C# | WebQueryConnection | external data connections | list web queries | batch workbook processing | Excel data connections | detect web query URLs | enumerate external connections | Aspose.Cells .NET
// Common Searches: list WebQuery connections Aspose.Cells C# | how to read external data connections from Excel using Aspose | batch detect web query URLs in multiple workbooks | enumerate WebQueryConnection objects in .NET | Aspose.Cells get connection name and URL
// Developer Intent: Extract and display all WebQuery external connections from a set of Excel files.
// Use Cases: Generate an inventory of web‑query URLs across a portfolio of workbooks for governance audits. | Verify that no disallowed web queries are embedded before distributing Excel files to end users. | Log connection details to a central file to track external data sources referenced by corporate spreadsheets.
// AI Prompts: Create a method that returns a collection of WebQueryConnection objects from a Workbook using Aspose.Cells. | Extend the sample to write each detected WebQuery name, URL and IsHtmlTables flag to a CSV report. | Write code that removes every WebQueryConnection from a workbook and saves the cleaned file with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// The example loads each workbook from a supplied list, reads its DataConnections collection, filters for WebQueryConnection objects, and writes the connection name, URL and IsHtmlTables flag to the console, reporting when no web queries are present.
class DetectWebQueryConnections
{
    static void Main()
    {
        // List of workbook file paths to process
        string[] workbookFiles = new string[]
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx",
            // add more file paths as needed
        };

        foreach (string filePath in workbookFiles)
        {
            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(filePath);

            Console.WriteLine($"Workbook: {filePath}");

            // Access the collection of external connections
            ExternalConnectionCollection connections = workbook.DataConnections;

            bool foundWebQuery = false;

            // Iterate through all external connections
            for (int i = 0; i < connections.Count; i++)
            {
                ExternalConnection conn = connections[i];

                // Check if the connection is a WebQueryConnection
                if (conn is WebQueryConnection webQuery)
                {
                    foundWebQuery = true;
                    Console.WriteLine($"  WebQuery Connection #{i + 1}");
                    Console.WriteLine($"    Name : {webQuery.Name}");
                    Console.WriteLine($"    URL  : {webQuery.Url}");
                    Console.WriteLine($"    IsHtmlTables : {webQuery.IsHtmlTables}");
                }
            }

            if (!foundWebQuery)
            {
                Console.WriteLine("  No WebQuery connections found.");
            }

            // No modifications are made, but if a save were required, the following would apply:
            // workbook.Save(filePath); // uses the provided save rule
        }
    }
}
