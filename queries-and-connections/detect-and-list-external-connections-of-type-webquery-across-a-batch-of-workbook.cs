using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsWebQueryDetector
{
    class Program
    {
        static void Main(string[] args)
        {
            // List of workbook file paths to process.
            // Replace these paths with actual file locations.
            List<string> workbookFiles = new List<string>
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            foreach (string filePath in workbookFiles)
            {
                try
                {
                    // Load the workbook from file (lifecycle: load)
                    Workbook workbook = new Workbook(filePath);

                    Console.WriteLine($"Processing workbook: {filePath}");

                    // Access the collection of external connections.
                    ExternalConnectionCollection connections = workbook.DataConnections;

                    bool foundWebQuery = false;

                    // Iterate through all external connections.
                    for (int i = 0; i < connections.Count; i++)
                    {
                        ExternalConnection conn = connections[i];

                        // Check if the connection is a WebQueryConnection.
                        if (conn is WebQueryConnection webQuery)
                        {
                            foundWebQuery = true;
                            Console.WriteLine($"  WebQuery Connection #{i + 1}:");
                            Console.WriteLine($"    Name               : {webQuery.Name}");
                            Console.WriteLine($"    URL                : {webQuery.Url}");
                            Console.WriteLine($"    IsHtmlTables       : {webQuery.IsHtmlTables}");
                            Console.WriteLine($"    IsTextDates        : {webQuery.IsTextDates}");
                            Console.WriteLine($"    HtmlFormat         : {webQuery.HtmlFormat}");
                            Console.WriteLine($"    RefreshOnLoad      : {webQuery.RefreshOnLoad}");
                            Console.WriteLine($"    SaveData           : {webQuery.SaveData}");
                        }
                    }

                    if (!foundWebQuery)
                    {
                        Console.WriteLine("  No WebQuery connections found in this workbook.");
                    }

                    // No modifications are made, but if you need to save changes, use the save lifecycle rule:
                    // workbook.Save("Updated_" + System.IO.Path.GetFileName(filePath));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing workbook '{filePath}': {ex.Message}");
                }

                Console.WriteLine(); // Blank line between workbooks
            }
        }
    }
}