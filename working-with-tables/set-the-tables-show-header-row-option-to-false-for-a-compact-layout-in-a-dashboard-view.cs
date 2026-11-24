// Title: Hide the header row of an Excel ListObject to create a compact table layout for a dashboard using Aspose.Cells for .NET
// Common Searches: how to hide the header row of an Excel table with Aspose.Cells in C# | Aspose.Cells C# remove ListObject header for a compact dashboard layout | set ShowHeaderRow false on a ListObject using Aspose.Cells .NET | create a compact Excel table without header for reporting dashboards with Aspose.Cells | save workbook after hiding table header row with Aspose.Cells C#
// Tags: Aspose.Cells ListObject header visibility | C# ShowHeaderRow property false | Excel table without header using Aspose.Cells | reporting dashboard table formatting Aspose | Workbook.Save after ListObject modification

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example loads a workbook, checks for a ListObject on the first worksheet, sets its ShowHeaderRow property to false to produce a compact table layout suitable for dashboards, and saves the updated file.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (worksheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables (ListObjects) found in the worksheet.");
                return;
            }

            // Access the first table
            ListObject table = worksheet.ListObjects[0];

            // Hide the header row for a compact layout
            table.ShowHeaderRow = false;

            // Optional: apply a compact style
            // table.TableStyleType = TableStyleType.Light1;

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
