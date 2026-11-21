// Title: Disable the totals row of an Excel table with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide totals row in Excel ListObject | remove totals row from Excel table using Aspose.Cells .NET | set ShowTotals false for worksheet table in C# | how to turn off totals row display in Aspose.Cells workbook | programmatically hide Excel table totals row with Aspose.Cells API
// Tags: Aspose.Cells ShowTotals property | C# hide Excel table totals row | disable worksheet ListObject totals row | Aspose.Cells table display settings | Excel table totals row removal .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads a workbook, checks the first worksheet for a ListObject, sets its ShowTotals property to false to hide the totals row, and saves the result, demonstrating how to disable the totals row display in an Excel table using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // If the worksheet contains at least one table (ListObject)
            if (worksheet.ListObjects.Count > 0)
            {
                ListObject table = worksheet.ListObjects[0];

                // Disable the display of the totals row
                table.ShowTotals = false;
            }
            else
            {
                Console.WriteLine("No tables found in the worksheet.");
            }

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
