// Title: How to add and configure a WebQueryConnection in an Aspose.Cells workbook using C#
// Common Searches: Aspose.Cells C# create external WebQuery connection for HTML tables | set IsHtmlTables and HtmlFormatHandlingType on WebQueryConnection in C# | save workbook with persisted web query connection using Aspose.Cells | initialize WebQueryConnection by indexing DataConnections collection in C# | configure WebQueryConnection to treat text dates as dates Aspose.Cells
// Tags: WebQueryConnection configuration Aspose.Cells C# | external web data connection Excel Aspose.Cells | HtmlFormatHandlingType All setting Aspose.Cells | persist web query in XLSX workbook Aspose | IsHtmlTables property usage Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsWebQueryDemo
{
    // The example demonstrates creating an empty Workbook, accessing its DataConnections collection, adding a WebQueryConnection via the collection count index, setting the target URL, enabling HTML table extraction, treating text dates as dates, configuring HtmlFormatHandlingType to All, assigning a description, and saving the workbook as an XLSX file where the web query connection is persisted.
    class Program
    {
        static void Main(string[] args)
        {
            // Path where the resulting workbook will be saved
            const string outputPath = "WebQueryConnectionDemo.xlsx";

            try
            {
                // Create a new empty workbook
                Workbook workbook = new Workbook();

                // Access the collection that holds external data connections
                ExternalConnectionCollection connections = workbook.DataConnections;

                // Add a new WebQueryConnection to the collection.
                // The collection creates a new connection object when accessed at the index equal to Count.
                WebQueryConnection webQuery = (WebQueryConnection)connections[connections.Count];

                // Configure the web query connection
                webQuery.Url = "https://example.com/data.html";   // target URL
                webQuery.IsHtmlTables = true;                    // retrieve data from HTML tables
                webQuery.IsTextDates = true;                     // treat text dates as dates
                webQuery.HtmlFormat = HtmlFormatHandlingType.All; // handle all HTML formats

                // Optional: give the connection a friendly description
                webQuery.ConnectionDescription = "Sample web query for demo purposes";

                // Save the workbook – the connection is persisted inside the file
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine($"Web query URL: {webQuery.Url}");
                Console.WriteLine($"IsHtmlTables: {webQuery.IsHtmlTables}");
                Console.WriteLine($"IsTextDates: {webQuery.IsTextDates}");
                Console.WriteLine($"HtmlFormat: {webQuery.HtmlFormat}");
            }
            catch (Exception ex)
            {
                // Report any errors that occur during creation or saving
                Console.WriteLine("An error occurred while creating the WebQuery connection:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
