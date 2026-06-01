using System;
using System.IO;
using System.Threading;
using System.Collections.Generic;
using Aspose.Cells;

class XmlMapRefreshScheduler
{
    // List of workbook file paths to process
    private static readonly List<string> WorkbookPaths = new List<string>
    {
        @"C:\Workbooks\Report1.xlsx",
        @"C:\Workbooks\Report2.xlsx",
        // add more workbook paths as needed
    };

    static void Main()
    {
        // Calculate time until next midnight
        DateTime now = DateTime.Now;
        DateTime nextMidnight = now.Date.AddDays(1);
        TimeSpan dueTime = nextMidnight - now;

        // Set up a timer to run the refresh task daily at midnight
        Timer timer = new Timer(_ => RefreshAllWorkbooks(), null, dueTime, TimeSpan.FromDays(1));

        Console.WriteLine("XML map refresh scheduler started. Next run at: " + nextMidnight);
        Console.WriteLine("Press Enter to exit...");
        Console.ReadLine();

        // Clean up
        timer.Dispose();
    }

    private static void RefreshAllWorkbooks()
    {
        Console.WriteLine($"Refresh started at {DateTime.Now}");

        foreach (string path in WorkbookPaths)
        {
            try
            {
                RefreshWorkbookXmlMaps(path);
                Console.WriteLine($"Successfully refreshed: {path}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing {path}: {ex.Message}");
            }
        }

        Console.WriteLine($"Refresh completed at {DateTime.Now}");
    }

    private static void RefreshWorkbookXmlMaps(string workbookPath)
    {
        // Load the workbook
        using (Workbook wb = new Workbook(workbookPath))
        {
            // Refresh all connections, pivot tables, and charts (includes any linked XML data)
            wb.Worksheets.RefreshAll();

            // Iterate through each XML map and perform an export to ensure the map data is up‑to‑date.
            // The export is written to a temporary file and then discarded.
            foreach (XmlMap map in wb.Worksheets.XmlMaps)
            {
                string tempFile = Path.Combine(Path.GetTempPath(),
                    $"{Path.GetFileNameWithoutExtension(workbookPath)}_{map.Name}_{Guid.NewGuid()}.xml");

                // Export the XML map; this forces Aspose.Cells to re‑evaluate the map data.
                wb.ExportXml(map.Name, tempFile);

                // Delete the temporary file; we only needed the export to trigger the refresh.
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }

            // Save the workbook, overwriting the original file.
            wb.Save(workbookPath);
        }
    }
}