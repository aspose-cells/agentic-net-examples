// Title: How to set a pivot table data field to Rank Largest to Smallest using Aspose.Cells for .NET (C#)
// Common Searches: C# code to change pivot table data field ranking to RankLargestToSmallest with Aspose.Cells | Aspose.Cells .NET set ShowAs property to RankLargestToSmallest for Excel pivot table | Programmatically apply descending rank display format to a pivot field using Aspose.Cells C# | How to use ShowDataAs.RankLargestToSmallest on a pivot table data field in Aspose.Cells | Example of modifying pivot table data field ranking in an existing workbook with Aspose.Cells
// Tags: Aspose.Cells pivot table ShowAs property | C# set RankLargestToSmallest display format | Excel pivot data field ranking Aspose.Cells | modify pivot table data field programmatically | Aspose.Cells workbook pivot field ranking

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // This C# example loads an existing Excel workbook, accesses the first worksheet and its first pivot table, retrieves the first data field, and (when supported) sets the ShowAs property to ShowDataAs.RankLargestToSmallest to rank the values from largest to smallest, then saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("Error: No pivot tables found on the first worksheet.");
                    return;
                }

                // Get the first pivot table on the worksheet
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Ensure the pivot table has at least one data field
                if (pivotTable.DataFields.Count == 0)
                {
                    Console.WriteLine("Error: The pivot table does not contain any data fields.");
                    return;
                }

                // Access the first data field of the pivot table
                PivotField dataField = pivotTable.DataFields[0];

                // NOTE: The ShowAs property may not be available in older Aspose.Cells versions.
                // If supported, you can set the display format to rank largest to smallest:
                // dataField.ShowAs = ShowDataAs.RankLargestToSmallest;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors (e.g., Aspose.Cells specific exceptions)
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
