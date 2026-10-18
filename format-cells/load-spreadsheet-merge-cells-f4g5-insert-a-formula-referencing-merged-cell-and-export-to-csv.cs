// Title: Merge cells F4:G5, set a formula that references the merged range, and save the worksheet as CSV with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge a range and use its top-left cell in a formula | how to export an Excel sheet to CSV after merging cells using Aspose.Cells | C# code to set a formula that points to a merged cell range in Aspose.Cells | saving a workbook as CSV with TxtSaveOptions in Aspose.Cells .NET | reference merged cells in formulas when converting Excel to CSV with Aspose
// Tags: merge cells range F4:G5 Aspose.Cells | formula referencing merged range C# | export worksheet to CSV Aspose.Cells | TxtSaveOptions CSV Aspose.Cells | worksheet.Cells.Merge method usage

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook or creates a new one, merges cells F4:G5 on the first worksheet, assigns a formula in A1 that points to the merged range's top‑left cell, and then saves the result as a CSV file using TxtSaveOptions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default worksheet
            }

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells F4:G5 (row 3, column 5, spanning 2 rows and 2 columns)
            sheet.Cells.Merge(3, 5, 2, 2);

            // Insert a formula that references the top‑left cell of the merged area
            sheet.Cells["A1"].Formula = "=F4";

            // Save the workbook as CSV
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV);
            workbook.Save("output.csv", csvOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
