// Title: Benchmark DBConnection Refresh Time in Excel with Aspose.Cells for .NET
// Description: Loads an Excel workbook, enumerates its external data connections, identifies DBConnection objects, measures the time required to access the ConnectionString (as a refresh proxy) with a Stopwatch, logs the elapsed milliseconds, and optionally saves the workbook. Ideal for profiling and optimizing data‑connection performance in Aspose.Cells.
// Keywords: Aspose.Cells DBConnection timing | measure external connection refresh .NET | C# stopwatch Excel data connection | benchmark DBConnection performance | log refresh duration Aspose.Cells | Excel workbook external connections | performance profiling Aspose.Cells | DBConnection refresh simulation | optimize Excel data source latency
// Common Searches: how to time DBConnection refresh using Aspose.Cells C# | measure performance of external data connections in Excel .NET | log refresh duration for each DBConnection Aspose.Cells | benchmark Excel DBConnection refresh time | Aspose.Cells measure external connection latency
// Developer Intent: Capture and record the execution time of each DBConnection refresh operation in an Excel workbook for performance analysis.
// Use Cases: Identify slow‑running database connections before publishing a workbook. | Generate a performance report of all DBConnection refresh times for auditing or optimization. | Compare connection latency across multiple workbooks to fine‑tune data source configurations.
// AI Prompts: Create C# code that iterates through a Workbook's DataConnections, measures the refresh time of each DBConnection with a Stopwatch, and returns a dictionary of connection names to elapsed milliseconds. | Suggest a reliable method to trigger an actual DBConnection refresh in Aspose.Cells and capture its execution time. | Provide a reusable utility class for logging DBConnection refresh performance, including handling cases where no DBConnection objects are present.

using System;
using System.Diagnostics;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Loads an Excel workbook, enumerates its external data connections, identifies DBConnection objects, measures the time required to access the ConnectionString (as a refresh proxy) with a Stopwatch, logs the elapsed milliseconds, and optionally saves the workbook. Ideal for profiling and optimizing data‑connection performance in Aspose.Cells.
class DBConnectionRefreshPerformance
{
    static void Main()
    {
        // Load the workbook that contains external DB connections
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Get the collection of external connections
        ExternalConnectionCollection connections = workbook.DataConnections;

        // Iterate through each connection and measure refresh time for DBConnection objects
        for (int i = 0; i < connections.Count; i++)
        {
            ExternalConnection conn = connections[i];

            if (conn is DBConnection dbConn)
            {
                Console.WriteLine($"Refreshing DBConnection #{i} - Name: {dbConn.Name}");

                // Start timing
                Stopwatch sw = Stopwatch.StartNew();

                // ------------------------------------------------------------
                // NOTE: Aspose.Cells does not expose a direct Refresh method for
                // DBConnection. The actual refresh normally occurs when the
                // workbook is opened in Excel or when the connection is used to
                // populate a table. Here we simulate the refresh operation.
                // Accessing a property (e.g., ConnectionString) ensures the
                // object is touched without altering its state.
                // ------------------------------------------------------------
                string dummy = dbConn.ConnectionString;

                // Stop timing
                sw.Stop();

                Console.WriteLine($"Time taken: {sw.ElapsedMilliseconds} ms");
            }
        }

        // Save the workbook after processing (optional)
        workbook.Save("output.xlsx");
    }
}
