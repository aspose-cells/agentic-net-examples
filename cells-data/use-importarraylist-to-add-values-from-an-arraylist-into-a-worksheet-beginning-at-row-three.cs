// Title: Import an ArrayList into a worksheet at row 3 using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# import ArrayList into Excel starting at row 3 | How to use Worksheet.Cells.ImportArrayList to add data from a .NET ArrayList at a specific row | C# code to vertically import a collection into an Excel sheet with Aspose.Cells | Insert list of strings into the third row of an Excel workbook using Aspose.Cells | Aspose.Cells ImportArrayList method example with row offset in C#
// Tags: ImportArrayList row offset Aspose.Cells | vertical collection import Excel C# | populate worksheet from ArrayList Aspose.Cells | start import at specific row Aspose.Cells | C# Aspose.Cells import data from .NET collection

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first worksheet, builds an ArrayList of strings, and uses worksheet.Cells.ImportArrayList to vertically insert the list beginning at row index 2 (the third row) column 0, then saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare an ArrayList with sample data
            ArrayList data = new ArrayList
            {
                "Item 1",
                "Item 2",
                "Item 3",
                "Item 4"
            };

            // Import the ArrayList into the worksheet starting at row 3 (zero‑based index 2), column 0
            // isVertical = true places each item in a new row (vertical import)
            worksheet.Cells.ImportArrayList(data, 2, 0, true);

            // Define output file path
            string outputPath = "Result.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
