// Title: Scan a folder of .xlsx workbooks and list WebQuery external connections using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that iterates over all .xlsx files in a given directory, loads each workbook with Aspose.Cells, and prints the Name, Url, and IsHtmlTables properties of every WebQueryConnection found. | Modify the program to export the detected WebQuery connection details (workbook name, connection name, URL, IsHtmlTables) to a CSV report instead of writing to the console. | Add try‑catch logic that logs unreadable or corrupted workbook paths to a log file and continues processing the remaining files without terminating the batch.
// Common Searches: aspnet list web query connections in multiple excel files with Aspose.Cells | C# batch detect external data connections of type WebQuery in .xlsx workbooks | how to extract URL from WebQueryConnection using Aspose.Cells .NET | enumerate DataConnections and filter WebQueryConnection in a folder of Excel files | Aspose.Cells script to report HTML table web queries across many workbooks
// Tags: enumerate web query connections Aspose.Cells | batch process Excel .xlsx data connections | extract web query URLs from workbooks | filter DataConnections by web query type C# | report HTML table web queries Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// The example loads each .xlsx file in a specified folder with Aspose.Cells, iterates the workbook’s DataConnections collection, identifies WebQueryConnection objects, and writes their Name, Url, and IsHtmlTables values to the console (or optionally to a CSV). It reports workbooks with no WebQuery connections and disposes each workbook after processing.
class DetectWebQueryConnections
{
    static void Main()
    {
        // Folder containing the workbooks to be examined
        string folderPath = @"C:\Workbooks";

        // Retrieve all Excel files in the folder
        string[] workbookFiles = Directory.GetFiles(folderPath, "*.xlsx");

        foreach (string filePath in workbookFiles)
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(filePath);
            Console.WriteLine($"Workbook: {Path.GetFileName(filePath)}");

            bool hasWebQuery = false;

            // Access the external connections collection
            ExternalConnectionCollection connections = workbook.DataConnections;

            // Iterate through each connection
            for (int i = 0; i < connections.Count; i++)
            {
                ExternalConnection conn = connections[i];

                // Identify WebQuery connections
                if (conn is WebQueryConnection webConn)
                {
                    hasWebQuery = true;
                    Console.WriteLine($"  Connection #{i + 1}:");
                    Console.WriteLine($"    Name: {webConn.Name}");
                    Console.WriteLine($"    URL: {webConn.Url}");
                    Console.WriteLine($"    IsHtmlTables: {webConn.IsHtmlTables}");
                }
            }

            if (!hasWebQuery)
            {
                Console.WriteLine("  No WebQuery connections found.");
            }

            // Dispose the workbook (optional cleanup)
            workbook.Dispose();
        }
    }
}
