// Title: Hide an Excel table’s header row for a compact dashboard layout using Aspose.Cells in C#
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, accesses the first ListObject, sets ShowHeaderRow to false, and saves the file. | Generate a program that removes the header row from an Excel table to produce a tighter view in a dashboard using Aspose.Cells for .NET. | Create a script that checks for tables in a worksheet and disables their header rows for a compact presentation with Aspose.Cells.
// Common Searches: aspnet hide excel table header row Aspose.Cells | how to remove table header for dashboard view using Aspose.Cells C# | set ShowHeaderRow false Aspose.Cells example | compact Excel table layout without header Aspose.Cells .NET | disable ListObject header row programmatically Aspose.Cells
// Tags: Aspose.Cells ListObject ShowHeaderRow false | C# hide Excel table header using Aspose | compact dashboard table layout Aspose.Cells | modify Excel table properties programmatically .NET | disable table header row in workbook Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example loads an existing workbook, verifies that a table (ListObject) exists on the first worksheet, sets its ShowHeaderRow property to false to hide the header row for a more compact dashboard appearance, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count > 0)
            {
                // Access the first table in the worksheet
                ListObject table = sheet.ListObjects[0];

                // Hide the header row for a compact layout
                table.ShowHeaderRow = false;

                // Style adjustments are optional and may not be supported in older versions.
                // If needed, they can be applied using the TableStyleInfo property in newer releases.
            }
            else
            {
                Console.WriteLine("No tables found in the worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
