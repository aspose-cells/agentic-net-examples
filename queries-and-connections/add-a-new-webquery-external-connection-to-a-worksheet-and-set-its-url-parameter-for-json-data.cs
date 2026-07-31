// Title: Add a WebQueryConnection (JSON) to an Aspose.Cells Workbook in C#
// Description: Shows how to create a new Workbook, insert a WebQueryConnection that targets a JSON REST endpoint, set it as a non‑HTML/XML source, and save the workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | WebQueryConnection | JSON data source | external connection | C# .NET | add WebQuery | REST API workbook | configure external data | save workbook with connection | programmatic data refresh
// Common Searches: Aspose.Cells add WebQueryConnection C# | set JSON URL for WebQueryConnection .NET | external data connection JSON Aspose.Cells example | how to save workbook with WebQueryConnection | configure WebQueryConnection as JSON source
// Developer Intent: Create and configure a WebQueryConnection that points to a JSON URL, then save the workbook.
// Use Cases: Build a reporting workbook that pulls live JSON data from a REST API on demand. | Distribute a template with a pre‑configured WebQueryConnection so end users can refresh data without writing code. | Automate periodic data imports by programmatically adding multiple JSON‑based WebQueryConnections to a single workbook.
// AI Prompts: Generate C# code with Aspose.Cells that adds a WebQueryConnection to a workbook and retrieves JSON from a specified URL. | Explain how to set WebQueryConnection.IsHtmlTables and IsXml to false so the source is treated as JSON. | Provide a step‑by‑step guide to add several WebQueryConnections, each with a different JSON endpoint, and save the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExamples
{
    // Shows how to create a new Workbook, insert a WebQueryConnection that targets a JSON REST endpoint, set it as a non‑HTML/XML source, and save the workbook with Aspose.Cells for .NET.
    public class AddWebQueryConnectionDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new empty workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (optional, just to have a sheet in the file)
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a new WebQueryConnection to the workbook's DataConnections collection.
            // The collection creates a new connection when accessed with the current count index.
            WebQueryConnection webQuery = (WebQueryConnection)workbook.DataConnections[workbook.DataConnections.Count];

            // Set the URL that returns JSON data.
            // Example URL – replace with the actual endpoint as needed.
            webQuery.Url = "https://api.example.com/data.json";

            // Configure the connection: source is JSON, not HTML tables or XML.
            webQuery.IsHtmlTables = false;
            webQuery.IsXml = false;

            // Save the workbook with the new external connection.
            workbook.Save("WebQueryConnectionDemo.xlsx");
        }
    }
}
