// Title: How to hide the totals row of the first Excel table (ListObject) with Aspose.Cells for .NET
// AI Prompts: Load an existing workbook, retrieve the first ListObject on the first worksheet, assign false to its ShowTotals property, and save the updated file. | Programmatically deactivate the totals row for an Excel table by using Aspose.Cells in a C# application. | Update a .xlsx file so that the first table no longer displays a totals row, then write the changes to a new workbook.
// Common Searches: Aspose.Cells C# remove totals row from ListObject | Change ShowTotals setting for first ListObject using Aspose.Cells | Remove table totals row from worksheet with Aspose.Cells API | Turn off totals row when saving Excel file in C#
// Tags: Aspose.Cells hide table totals row | C# ListObject ShowTotals configuration | Excel table totals row removal Aspose | modify ListObject properties Aspose.Cells | disable totals row workbook C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an existing Excel workbook, checks for tables on the first worksheet, sets the ShowTotals property of the first ListObject to false to hide its totals row, and saves the modified workbook to a new file.
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
            Worksheet sheet = workbook.Worksheets[0];

            // Check if the worksheet contains any tables (ListObjects)
            if (sheet.ListObjects.Count > 0)
            {
                // Access the first table
                ListObject table = sheet.ListObjects[0];

                // Hide the totals row
                table.ShowTotals = false;
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
