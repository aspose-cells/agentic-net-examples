// Title: Hide an Excel table’s header row while preserving data rows using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an existing XLSX file, finds the first ListObject in the first worksheet, sets its ShowHeaderRow property to false, and saves the workbook to a new file using Aspose.Cells. | Show how to verify that a worksheet contains at least one ListObject before attempting to change its header visibility with Aspose.Cells. | Describe the process of disabling a table’s header row while leaving all data rows intact in a .NET application using the Aspose.Cells API.
// Common Searches: asp.net hide excel table header row using aspose.cells | c# hide listobject header without removing rows | aspose.cells set ShowHeaderRow false for first table | programmatically hide excel table header while keeping data rows | how to suppress table header in workbook with aspose cells c#
// Tags: Aspose.Cells ListObject ShowHeaderRow property | Excel table header visibility .NET | disable table header without data loss Aspose.Cells | C# workbook table header manipulation | preserve rows while hiding table header

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example loads 'input.xlsx', checks whether the first worksheet contains any ListObjects, disables the header row of the first table by setting ShowHeaderRow = false, and saves the result as 'output.xlsx', handling missing files and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // If the worksheet contains tables (list objects), hide the header row of the first table
                if (worksheet.ListObjects.Count > 0)
                {
                    ListObject table = worksheet.ListObjects[0];
                    table.ShowHeaderRow = false;
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
