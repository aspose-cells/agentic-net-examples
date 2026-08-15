// Title: Get PivotTable Last Refresh Date with Aspose.Cells in C# (.NET)
// Description: This example demonstrates how to load an Excel workbook using Aspose.Cells for .NET, locate the first PivotTable on the first worksheet, read its RefreshDate property to obtain the timestamp of the most recent refresh, and display the table name and refresh date in the console. Includes error handling for missing files and worksheets without PivotTables.
// Keywords: Aspose.Cells | C# | PivotTable RefreshDate | last refresh timestamp | read pivot metadata | load Excel workbook | retrieve pivot refresh date | Excel pivot table date | Aspose.Cells example | .NET
// Common Searches: Aspose.Cells get PivotTable refresh date C# | How to read PivotTable RefreshDate with Aspose.Cells | C# code to obtain last refresh time of an Excel PivotTable | Retrieve PivotTable last refreshed timestamp using Aspose.Cells | Example of reading PivotTable metadata in .NET
// Developer Intent: Load an Excel file, locate a PivotTable, and extract its RefreshDate value.
// Use Cases: Show the most recent refresh time of a PivotTable on a reporting dashboard. | Validate that a PivotTable has been refreshed after a data load before generating downstream reports. | Log refresh timestamps of all PivotTables in a workbook for audit or compliance purposes.
// AI Prompts: Write C# code with Aspose.Cells that iterates through every PivotTable in a workbook, records each table's RefreshDate, and writes the results to a CSV log. | Create a snippet that checks if a PivotTable's RefreshDate is older than a given date and, if so, triggers a refresh using Aspose.Cells. | Explain best practices for safely handling worksheets that contain no PivotTables when accessing the RefreshDate property.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshDateDemo
{
    // This example demonstrates how to load an Excel workbook using Aspose.Cells for .NET, locate the first PivotTable on the first worksheet, read its RefreshDate property to obtain the timestamp of the most recent refresh, and display the table name and refresh date in the console. Includes error handling for missing files and worksheets without PivotTables.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file that contains a PivotTable
            string inputPath = "PivotTableSample.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one PivotTable
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No PivotTables found in the worksheet.");
                    return;
                }

                // Retrieve the first PivotTable in the collection
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Read the RefreshDate property which indicates the last refresh time
                DateTime refreshDate = pivotTable.RefreshDate;

                // Output the refresh date information
                Console.WriteLine($"Pivot Table Name: {pivotTable.Name}");
                Console.WriteLine($"Last Refresh Date: {refreshDate}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
