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
            // List of workbook file paths to process
            string[] workbookFiles = new string[]
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            foreach (string filePath in workbookFiles)
            {
                // Load the workbook (lifecycle: load)
                Workbook workbook = new Workbook(filePath);

                Console.WriteLine($"Processing workbook: {filePath}");

                // Access the collection of external connections
                ExternalConnectionCollection connections = workbook.DataConnections;

                bool hasWebQuery = false;

                // Iterate through all external connections
                for (int i = 0; i < connections.Count; i++)
                {
                    ExternalConnection conn = connections[i];

                    // Check if the connection is a WebQueryConnection
                    if (conn is WebQueryConnection webQuery)
                    {
                        hasWebQuery = true;
                        Console.WriteLine($"  WebQuery Connection #{i + 1}");
                        Console.WriteLine($"    Name : {webQuery.Name}");
                        Console.WriteLine($"    URL  : {webQuery.Url}");
                        Console.WriteLine($"    IsHtmlTables : {webQuery.IsHtmlTables}");
                        Console.WriteLine($"    HtmlFormat   : {webQuery.HtmlFormat}");
                    }
                }

                if (!hasWebQuery)
                {
                    Console.WriteLine("  No WebQuery connections found.");
                }

                // No modifications are made, so no need to save the workbook.
                // If you wanted to persist changes, you would use:
                // workbook.Save("Updated_" + System.IO.Path.GetFileName(filePath));
            }

            Console.WriteLine("Detection completed.");
        }
    }
}