// Title: Enable RefreshOnLoad for a WebQueryConnection in Excel using Aspose.Cells for .NET
// Description: Demonstrates loading an existing .xlsx file, locating a WebQueryConnection, setting its RefreshOnLoad property to true for automatic data refresh on workbook open, and saving the updated file.
// Keywords: Aspose.Cells | WebQueryConnection | RefreshOnLoad | C# .NET | Excel data connection | automatic refresh | set RefreshOnLoad true | update web query | Excel workbook load | Aspose.Cells API
// Common Searches: Aspose.Cells set RefreshOnLoad | How to make web query refresh on workbook open C# | Enable automatic refresh for WebQueryConnection | RefreshOnLoad property Aspose.Cells example | Update Excel web query connection programmatically
// Developer Intent: Set the RefreshOnLoad flag so a WebQueryConnection updates its data automatically whenever the workbook is opened.
// Use Cases: Programmatically guarantee that external web data is refreshed each time an Excel file is opened. | Batch‑process multiple workbooks to enable auto‑refresh of web queries before distribution. | Validate the existence of a WebQueryConnection before modifying its RefreshOnLoad setting to avoid runtime errors. | Integrate automatic web‑query refresh into data‑pipeline workflows that rely on live Excel reports.
// AI Prompts: Generate C# code with Aspose.Cells that iterates over all WebQueryConnection objects in a workbook and sets RefreshOnLoad = true. | Provide a try‑catch example that loads a workbook, checks for a WebQueryConnection, toggles RefreshOnLoad, logs the result, and saves the file. | Explain how the RefreshOnLoad property interacts with Excel's connection settings and how to disable it when needed.

using Aspose.Cells;
using Aspose.Cells.ExternalConnections;
using System;
using System.IO;

// Demonstrates loading an existing .xlsx file, locating a WebQueryConnection, setting its RefreshOnLoad property to true for automatic data refresh on workbook open, and saving the updated file.
class SetRefreshOnLoadDemo
{
    static void Main()
    {
        const string inputPath = "WebQuerySample.xlsx";
        const string outputPath = "WebQuerySample_RefreshOnLoad.xlsx";

        try
        {
            // Ensure the input workbook exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains a WebQuery connection.
            Workbook workbook = new Workbook(inputPath);

            // Check for a WebQueryConnection and set RefreshOnLoad.
            if (workbook.DataConnections.Count > 0 && workbook.DataConnections[0] is WebQueryConnection webConn)
            {
                webConn.RefreshOnLoad = true;
                Console.WriteLine($"RefreshOnLoad set to: {webConn.RefreshOnLoad}");
            }
            else
            {
                Console.WriteLine("No WebQueryConnection found in the workbook.");
            }

            // Save the workbook with updated settings.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
