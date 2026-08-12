// Title: Nightly C# scheduler to refresh XML maps in Excel workbooks using Aspose.Cells
// Description: A C# console app that builds a list of workbook paths, calculates the interval to the next midnight, and uses System.Threading.Timer to invoke a RefreshAllXmlMaps method every 24 hours. The method loads each workbook with Aspose.Cells, calls Worksheets.RefreshAll() to update all XML maps, saves the file, and writes success or error messages to the console.
// Keywords: Aspose.Cells C# XML map refresh | daily Excel workbook refresh | System.Threading.Timer schedule | midnight task Aspose | automate XML map update | refresh all worksheets | batch Excel processing | Windows service Excel refresh
// Common Searches: C# schedule task to refresh XML maps in Excel | Aspose.Cells refresh all worksheets nightly | How to automate XML map refresh at midnight | Batch refresh XML maps with Aspose.Cells | Create Windows service for daily Excel refresh
// Developer Intent: Create an automated nightly routine that loads each listed workbook, refreshes its XML maps, saves the changes, and logs the outcome.
// Use Cases: Keep financial reporting workbooks up‑to‑date each night before morning analysis. | Update XML maps in a batch of client spreadsheets after a nightly data import. | Run a background service that guarantees all stored Excel templates have current XML map connections before user access.
// AI Prompts: Generate C# code that uses Aspose.Cells and System.Threading.Timer to run a midnight job that refreshes XML maps in a list of workbook files and logs results. | Show how to handle exceptions per workbook while refreshing XML maps and saving with Aspose.Cells in a scheduled task. | Explain how to convert the console timer into a Windows Service or Azure Function so the nightly XML map refresh persists after application restart.

using System;
using System.Collections.Generic;
using System.Threading;
using Aspose.Cells;

// A C# console app that builds a list of workbook paths, calculates the interval to the next midnight, and uses System.Threading.Timer to invoke a RefreshAllXmlMaps method every 24 hours. The method loads each workbook with Aspose.Cells, calls Worksheets.RefreshAll() to update all XML maps, saves the file, and writes success or error messages to the console.
class Program
{
    // List of workbook file paths to be refreshed
    static readonly List<string> workbookPaths = new List<string>
    {
        @"C:\Workbooks\Book1.xlsx",
        @"C:\Workbooks\Book2.xlsx"
        // add more paths as needed
    };

    static void Main()
    {
        // Calculate the interval until the next midnight
        DateTime now = DateTime.Now;
        DateTime nextMidnight = now.Date.AddDays(1);
        TimeSpan timeToMidnight = nextMidnight - now;

        // Set up a timer that triggers at midnight and then every 24 hours
        Timer timer = new Timer(
            callback: state => RefreshAllXmlMaps(),
            state: null,
            dueTime: timeToMidnight,
            period: TimeSpan.FromDays(1));

        // Keep the application running
        Console.WriteLine("XML map refresh scheduler started. Press Enter to exit.");
        Console.ReadLine();
    }

    // Refreshes XML maps in all specified workbooks
    static void RefreshAllXmlMaps()
    {
        foreach (string path in workbookPaths)
        {
            try
            {
                // Load the workbook (uses the load rule)
                Workbook wb = new Workbook(path);

                // Refresh all connections/pivot tables (covers XML map refresh)
                wb.Worksheets.RefreshAll();

                // Save the workbook back to the same file (uses the save rule)
                wb.Save(path);

                Console.WriteLine($"Successfully refreshed XML maps in '{path}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{path}': {ex.Message}");
            }
        }
    }
}
