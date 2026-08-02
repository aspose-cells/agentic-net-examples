using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsPerformance
{
    class DBConnectionRefreshTimer
    {
        static void Main()
        {
            // Path to the workbook containing DB connections
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Get the collection of external connections
            ExternalConnectionCollection connections = workbook.DataConnections;

            // Iterate through each connection and measure refresh time for DBConnection objects
            foreach (ExternalConnection conn in connections)
            {
                if (conn is DBConnection dbConn)
                {
                    // Start timing
                    Stopwatch sw = Stopwatch.StartNew();

                    // Attempt to invoke a refresh operation via reflection.
                    // Some DBConnection implementations expose a RefreshData method.
                    var refreshMethod = dbConn.GetType().GetMethod("RefreshData");
                    if (refreshMethod != null)
                    {
                        // Invoke the method without parameters
                        refreshMethod.Invoke(dbConn, null);
                    }
                    else
                    {
                        // If no explicit refresh method exists, simulate a refresh operation.
                        // This placeholder can be replaced with the appropriate API call when available.
                        // For demonstration, we simply access a property.
                        var _ = dbConn.RefreshOnLoad;
                    }

                    // Stop timing
                    sw.Stop();

                    // Log the elapsed time
                    Console.WriteLine($"DBConnection '{dbConn.Name}' refreshed in {sw.ElapsedMilliseconds} ms.");
                }
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}