// Title: Copy a range of cells containing formulas to a new location while preserving calculation dependencies with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to copy a source range that includes formulas to a destination range and have the library automatically adjust the cell references. | Load or create a workbook, define source and target ranges, and call the Range.Copy method to duplicate formulas while keeping their calculation relationships intact.
// Common Searches: Aspose.Cells copy range with formulas and keep relative references in C# | How to duplicate a cell block containing formulas to another area using Aspose.Cells for .NET | Preserve formula dependencies when moving a range in an Excel workbook with Aspose.Cells | C# Aspose.Cells Range.Copy updates cell references automatically
// Tags: Aspose.Cells range copy with formulas | preserve formula references Aspose.Cells C# | copy cell block maintaining dependencies .NET | Range.Copy method Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook or creates a new one with sample data and formulas in A1:C5, defines a source range (A1:C5) and a destination range (E1:G5), copies the source range using Range.Copy which preserves the formulas and automatically updates cell references, and saves the result as result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string resultPath = "result.xlsx";

            // Ensure the source file exists; if not, create a new workbook with sample data.
            Workbook workbook;
            if (File.Exists(sourcePath))
            {
                workbook = new Workbook(sourcePath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                // Populate sample data and formulas in A1:C5
                ws.Cells["A1"].PutValue(1);
                ws.Cells["B1"].PutValue(2);
                ws.Cells["C1"].Formula = "=A1+B1";
                ws.Cells["A2"].PutValue(3);
                ws.Cells["B2"].PutValue(4);
                ws.Cells["C2"].Formula = "=A2+B2";
                // Extend as needed...
                workbook.Save(sourcePath);
            }

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the source range that contains formulas
            Aspose.Cells.Range sourceRange = sheet.Cells.CreateRange("A1:C5");

            // Define the destination range where the formulas will be copied
            Aspose.Cells.Range destinationRange = sheet.Cells.CreateRange("E1:G5");

            // Copy the source range to the destination.
            // This preserves formulas and automatically updates cell references.
            sourceRange.Copy(destinationRange);

            // Save the workbook with the updated data
            workbook.Save(resultPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
