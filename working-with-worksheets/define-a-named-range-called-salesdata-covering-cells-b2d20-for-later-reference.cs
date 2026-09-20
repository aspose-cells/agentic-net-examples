// Title: Define a named range "SalesData" for cells B2:D20 in a new workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, add a named range called SalesData that refers to B2:D20 on the first worksheet, and save it as Output.xlsx. | Write C# code that builds the address string for B2:D20, assigns it to the RefersTo property of a named range named SalesData, and writes the workbook to disk.
// Common Searches: Aspose.Cells C# how to add a named range that points to B2:D20 | C# set RefersTo property for a named range in Aspose.Cells workbook | Create named range SalesData in Aspose.Cells and export to Excel file
// Tags: Aspose.Cells define named range B2:D20 | C# add named range SalesData | Aspose.Cells set RefersTo property | C# create workbook with named range | Aspose.Cells named range worksheet reference

using System;
using System.IO;
using Aspose.Cells;

// // This example creates a new workbook, defines a named range called SalesData that references cells B2:D20 on the first worksheet, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range B2:D20 (zero‑based indices)
            int firstRow = 1;      // Row 2
            int firstColumn = 1;   // Column B
            int totalRows = 19;    // Rows 2‑20 inclusive
            int totalColumns = 3;  // Columns B‑D inclusive

            // Build the address string for the range (e.g., 'Sheet1'!$B$2:$D$20)
            string startCell = CellsHelper.CellIndexToName(firstRow, firstColumn);
            string endCell = CellsHelper.CellIndexToName(firstRow + totalRows - 1, firstColumn + totalColumns - 1);
            string rangeAddress = $"='{sheet.Name}'!${startCell}:${endCell}";

            // Add a named range called "SalesData"
            workbook.Worksheets.Names.Add("SalesData");
            int nameIndex = workbook.Worksheets.Names.Count - 1;
            workbook.Worksheets.Names[nameIndex].RefersTo = rangeAddress;

            // Define output file path
            string outputPath = "Output.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
