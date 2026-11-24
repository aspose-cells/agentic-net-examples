// Title: Copy a named range between worksheets while preserving absolute cell references using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy named range to another sheet preserving absolute references | How to transfer a named range from Sheet1 to Sheet2 with formulas intact using Aspose.Cells | C# example using GetRangeByName and Range.Copy to move a named range in Aspose.Cells | Preserve absolute cell references when copying a named range between worksheets in Aspose.Cells | Create a destination range and copy a named range with Aspose.Cells .NET
// Tags: named range transfer Aspose.Cells | GetRangeByName usage C# | Range.Copy method preserving formulas | create destination range Aspose.Cells | absolute references preservation Aspose.Cells | worksheet range copy .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook, retrieves a named range with GetRangeByName, creates an equally sized range on a destination worksheet (adding the sheet if necessary), copies the source range using Range.Copy—which automatically retains absolute cell references in formulas—and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook that contains the named range
            Workbook workbook = new Workbook(inputPath);

            // Source worksheet where the named range is defined
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"]; // adjust sheet name as needed
            if (sourceSheet == null)
                throw new InvalidOperationException("Source worksheet 'Sheet1' not found.");

            // Destination worksheet (create if it does not exist)
            Worksheet destinationSheet = workbook.Worksheets["Sheet2"];
            if (destinationSheet == null)
                destinationSheet = workbook.Worksheets.Add("Sheet2");

            // Retrieve the named range by its name (e.g., "MyRange")
            // Use the Workbook's GetRangeByName method (returns Aspose.Cells.Range)
            Aspose.Cells.Range sourceRange = workbook.Worksheets.GetRangeByName("MyRange");
            if (sourceRange == null)
                throw new InvalidOperationException("Named range 'MyRange' not found.");

            // Define the top‑left cell in the destination sheet where the range will be copied
            int destStartRow = 0;      // Row 0 corresponds to Excel row 1 (A1)
            int destStartColumn = 0;   // Column 0 corresponds to Excel column A

            // Create a range in the destination sheet with the same dimensions as the source range
            Aspose.Cells.Range destinationRange = destinationSheet.Cells.CreateRange(
                destStartRow,
                destStartColumn,
                sourceRange.RowCount,
                sourceRange.ColumnCount);

            // Copy the source range to the destination range.
            // Absolute cell references inside formulas are preserved automatically.
            destinationRange.Copy(sourceRange);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
