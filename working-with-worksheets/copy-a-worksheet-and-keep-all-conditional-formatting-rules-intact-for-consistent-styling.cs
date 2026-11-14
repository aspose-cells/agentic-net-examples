// Title: Copy an Excel worksheet and retain all conditional formatting using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy worksheet without losing conditional formatting | How to duplicate a sheet in an .xlsx file while keeping conditional format rules | Preserve conditional formatting when using AddCopy method in Aspose.Cells | Copy Excel worksheet programmatically in .NET and retain its styling | C# code to copy a worksheet and keep its conditional formatting intact
// Tags: copy worksheet with conditional formatting Aspose.Cells | AddCopy API retains conditional formats | duplicate Excel sheet preserving styling C# | Aspose.Cells worksheet copy sample | retain conditional formatting on sheet copy .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads a source .xlsx file, verifies its presence, selects a worksheet, creates a copy of that worksheet within the same workbook using the AddCopy method (which automatically carries over all conditional formatting rules), optionally renames the new sheet, and saves the workbook to a new file.
class WorksheetCopyExample
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(sourcePath);

            // Reference the worksheet to be copied (by name or index)
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"]; // replace with actual sheet name

            if (sourceSheet == null)
            {
                Console.WriteLine("Worksheet 'Sheet1' not found in the source workbook.");
                return;
            }

            // Copy the worksheet to a new position in the same workbook
            // AddCopy expects the worksheet name (string)
            int copiedIndex = workbook.Worksheets.AddCopy(sourceSheet.Name);

            // Optionally rename the copied worksheet
            workbook.Worksheets[copiedIndex].Name = "Sheet1_Copy";

            // Save the workbook with the copied worksheet
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
