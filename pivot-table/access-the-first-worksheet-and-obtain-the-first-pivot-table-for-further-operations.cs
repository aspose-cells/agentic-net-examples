// Title: How to retrieve the first pivot table from the first worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get first pivot table from worksheet | Aspose.Cells .NET retrieve pivot table by index example | Read pivot table name from first sheet using Aspose.Cells | Check if a worksheet contains pivot tables with Aspose.Cells C# | Load Excel file and access pivot tables collection Aspose.Cells
// Tags: Aspose.Cells workbook load and pivot table access | retrieve first pivot table Aspose.Cells C# | pivot table collection index Aspose.Cells | verify pivot table count on worksheet Aspose.Cells | output pivot table name Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    // The example loads an Excel workbook, accesses the first worksheet, verifies that it contains at least one pivot table, retrieves the first pivot table by its index, prints the pivot table's name, and includes error handling for missing files or runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (index 0)
                Worksheet firstWorksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (firstWorksheet.PivotTables.Count > 0)
                {
                    // Obtain the first pivot table (index 0) for further operations
                    PivotTable firstPivotTable = firstWorksheet.PivotTables[0];

                    // Example: output the name of the pivot table
                    Console.WriteLine("First Pivot Table Name: " + firstPivotTable.Name);
                }
                else
                {
                    Console.WriteLine("No pivot tables found on the first worksheet.");
                }
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
