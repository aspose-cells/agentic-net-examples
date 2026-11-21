// Title: Set a Pivot Table Data Field to RankLargestToSmallest Display Format with Aspose.Cells for .NET (C#)
// Common Searches: aspocells c# set pivot table data field ranking to descending | how to apply RankLargestToSmallest display format to a pivot field using Aspose.Cells | C# example for changing pivot table data field display format in Excel | programmatically set pivot table ranking order with Aspose.Cells for .NET | modify pivot table data field format to show largest values first in C#
// Tags: aspocells pivot field display format | c# set pivot table ranking order | excel pivot table descending rank via aspocells | modify pivot table data field format c# | aspocells .net pivot table formatting

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable and PivotField classes

namespace AsposeCellsExamples
{
    // Demonstrates loading an Excel workbook, accessing the first pivot table's first data field, and configuring its display format (e.g., RankLargestToSmallest for descending ranking) using Aspose.Cells for .NET before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                // Retrieve the first pivot table
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Ensure the pivot table has at least one data field
                if (pivotTable.DataFields.Count == 0)
                {
                    Console.WriteLine("Pivot table contains no data fields.");
                    return;
                }

                // Get the first data field
                PivotField dataField = pivotTable.DataFields[0];

                // Set a numeric format as an example (replace with desired format)
                dataField.NumberFormat = "#,##0";

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
