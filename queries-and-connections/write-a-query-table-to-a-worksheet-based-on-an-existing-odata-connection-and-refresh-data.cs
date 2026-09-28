// Title: Refresh an OData query table in an existing worksheet while preserving formatting using Aspose.Cells for .NET
// AI Prompts: Load a workbook that contains a WebQueryConnection, assign a new OData endpoint to its Url, disable BackgroundRefresh, enable RefreshOnLoad, keep existing cell styles, invoke Workbook.RefreshAll, and save the refreshed file. | Programmatically adjust the properties of the first query table's WebQueryConnection, trigger a synchronous OData refresh, ensure formatting is retained, and write the updated workbook to a new location.
// Common Searches: Aspose.Cells C# how to change OData service URL in a query table and refresh data | C# refresh external data connections in Excel workbook using Aspose.Cells | Preserve cell formatting when refreshing OData query tables with Aspose.Cells .NET | Update WebQueryConnection properties and call RefreshAll in Aspose.Cells example | Read an Excel file that already has an OData connection and programmatically refresh its query table
// Tags: modify WebQueryConnection URL Aspose.Cells | synchronous OData query table refresh Aspose.Cells | maintain cell styles during OData query refresh Aspose.Cells | refresh all external connections Aspose.Cells .NET | open workbook containing OData connection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsQueryTableODataDemo
{
    // The sample loads a workbook with an existing OData (Web) connection and query table, updates the connection URL and refresh settings, enables synchronous refresh, preserves cell formatting, calls RefreshAll to pull fresh data, and saves the refreshed workbook to a new file.
    class Program
    {
        static void Main()
        {
            const string templatePath = "TemplateWithODataConnection.xlsx";
            const string outputPath = "Output_WithRefreshedQueryTable.xlsx";

            try
            {
                // Verify that the template file exists before attempting to load it
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the workbook that already contains an OData (Web) connection and a query table
                Workbook workbook = new Workbook(templatePath);

                // Access the first worksheet where the query table should reside
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure that a query table exists in the worksheet
                if (worksheet.QueryTables.Count == 0)
                {
                    Console.WriteLine("No query tables found in the worksheet.");
                    return;
                }

                // Get the first query table
                QueryTable queryTable = worksheet.QueryTables[0];

                // Retrieve the external connection associated with the query table.
                // For OData connections Aspose.Cells uses the WebQueryConnection class.
                ExternalConnection externalConn = queryTable.ExternalConnection;
                if (externalConn is WebQueryConnection oDataConn)
                {
                    // Adjust connection properties before refresh
                    oDataConn.BackgroundRefresh = false;          // refresh synchronously
                    oDataConn.RefreshOnLoad = true;               // refresh when workbook loads
                    oDataConn.IsHtmlTables = false;               // not relevant for OData
                    oDataConn.Url = "https://services.odata.org/V4/Northwind/Northwind.svc/Products";
                }
                else
                {
                    Console.WriteLine("The external connection is not a WebQueryConnection (OData).");
                    return;
                }

                // Preserve existing cell formatting after refresh
                queryTable.PreserveFormatting = true;

                // Refresh all external data connections (including the query table)
                try
                {
                    workbook.RefreshAll();
                }
                catch (Exception refreshEx)
                {
                    Console.WriteLine($"Error during refresh: {refreshEx.Message}");
                    return;
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Query table refreshed and workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
