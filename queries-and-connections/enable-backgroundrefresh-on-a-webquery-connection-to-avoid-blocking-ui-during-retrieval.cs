// Title: Enable BackgroundRefresh for WebQueryConnection in Aspose.Cells (.NET) to Avoid UI Freeze
// Description: Demonstrates how to load an existing workbook, locate its WebQueryConnection, set the BackgroundRefresh property to true for asynchronous data retrieval, and save the workbook, preventing UI blocking in .NET applications.
// Keywords: Aspose.Cells | WebQueryConnection | BackgroundRefresh | .NET | C# | asynchronous web query | prevent UI freeze | data connection settings | Excel workbook automation
// Common Searches: Aspose.Cells enable BackgroundRefresh C# | WebQueryConnection async refresh .NET | prevent UI blocking Aspose.Cells web query | set BackgroundRefresh property programmatically | update Excel data connection with Aspose.Cells
// Developer Intent: Set BackgroundRefresh = true on a WebQueryConnection so the query runs in the background and does not block the UI.
// Use Cases: Load a workbook that contains a web query, enable background refresh, and save the updated file. | Iterate through all data connections in a workbook and activate BackgroundRefresh for each WebQueryConnection before publishing the workbook. | Improve responsiveness of a WinForms or WPF app that refreshes web‑based Excel data using Aspose.Cells.
// AI Prompts: Generate C# code with Aspose.Cells that enables BackgroundRefresh for every WebQueryConnection in a workbook and saves the changes. | Explain the impact of BackgroundRefresh on performance and UI behavior when using Aspose.Cells WebQuery connections. | Provide best‑practice error handling for loading a workbook and modifying its data connections with Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.ExternalConnections;
using System;
using System.IO;

// Demonstrates how to load an existing workbook, locate its WebQueryConnection, set the BackgroundRefresh property to true for asynchronous data retrieval, and save the workbook, preventing UI blocking in .NET applications.
class EnableBackgroundRefreshDemo
{
    static void Main()
    {
        const string inputPath = "InputWithWebQuery.xlsx";
        const string outputPath = "OutputWithBackgroundRefresh.xlsx";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load a workbook that already contains a WebQuery connection
            Workbook workbook = new Workbook(inputPath);

            // Locate the first WebQueryConnection in the workbook's data connections
            WebQueryConnection webConn = null;
            foreach (ExternalConnection conn in workbook.DataConnections)
            {
                if (conn is WebQueryConnection wqc)
                {
                    webConn = wqc;
                    break;
                }
            }

            if (webConn == null)
            {
                Console.WriteLine("No WebQueryConnection found in the workbook.");
                return;
            }

            // Enable background refresh to prevent UI blocking during data retrieval
            webConn.BackgroundRefresh = true;
            Console.WriteLine("BackgroundRefresh set to: " + webConn.BackgroundRefresh);

            // Save the workbook with the updated connection settings
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
