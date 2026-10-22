// Title: Hide the totals row of an Excel table using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, accesses the first ListObject in the first worksheet, sets its ShowTotals property to false, and saves the updated file. | Show how to programmatically disable the totals row of an Excel table in a .NET application by using the Aspose.Cells ListObject API.
// Common Searches: C# Aspose.Cells hide totals row in Excel table | How to set ShowTotals false for a ListObject with Aspose.Cells | Remove totals row from Excel worksheet programmatically using Aspose.Cells .NET | Aspose.Cells disable table totals row without affecting data | Set ShowTotals property to false for first table in workbook Aspose.Cells
// Tags: Aspose.Cells ListObject ShowTotals false | C# disable Excel table totals row | Aspose.Cells modify table display options | Excel table totals row removal .NET | Aspose.Cells suppress table totals row

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example loads Input.xlsx, checks for a ListObject in the first worksheet, sets its ShowTotals property to false to hide the totals row, and saves the result as Output.xlsx, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count > 0)
            {
                // Get the first table in the worksheet
                ListObject table = sheet.ListObjects[0];

                // Disable the display of the totals row
                table.ShowTotals = false;
            }
            else
            {
                Console.WriteLine("No tables found in the worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
