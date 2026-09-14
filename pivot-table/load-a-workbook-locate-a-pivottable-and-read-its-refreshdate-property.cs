// Title: How to read a PivotTable's RefreshDate property from an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, locates the first PivotTable, and prints its RefreshDate. | Provide a .NET snippet that verifies a workbook exists, accesses the PivotTable collection, and returns the last refresh timestamp for each PivotTable. | Create an error‑handled C# example that reads the RefreshDate of a specific PivotTable by name using Aspose.Cells.
// Common Searches: Aspose.Cells C# read pivot table last refreshed date from existing workbook | Get RefreshDate of PivotTable in .xlsx using Aspose.Cells .NET | How to retrieve pivot table refresh timestamp with Aspose.Cells library | C# code to access PivotTable.RefreshDate property in an Excel file
// Tags: aspocells read pivot refreshdate | c# aspocells pivot table metadata | load workbook retrieve pivot refresh timestamp | excel pivot refreshdate aspocells .net | aspocells get pivot last refreshed time

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example demonstrates loading an existing Excel workbook with Aspose.Cells, accessing the first worksheet's PivotTable collection, reading the RefreshDate property of the first PivotTable, and outputting the pivot name together with its last refresh timestamp, including file existence checks and exception handling.
    public class PivotTableRefreshDateReader
    {
        public static void Run()
        {
            // Path to the existing Excel file that contains a pivot table
            string inputPath = "PivotTableSample.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook from the file
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the collection of pivot tables in this worksheet
                PivotTableCollection pivotTables = worksheet.PivotTables;

                // Ensure there is at least one pivot table
                if (pivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                // Retrieve the first pivot table (or use a specific name/index as required)
                PivotTable pivotTable = pivotTables[0];

                // Read the RefreshDate property which indicates the last time the pivot table was refreshed
                DateTime refreshDate = pivotTable.RefreshDate;

                // Output the refresh date
                Console.WriteLine($"Pivot Table \"{pivotTable.Name}\" Refresh Date: {refreshDate}");
            }
            catch (Exception ex)
            {
                // Handle any errors that occur during processing
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }

    public static class Program
    {
        public static void Main(string[] args)
        {
            PivotTableRefreshDateReader.Run();
        }
    }
}
