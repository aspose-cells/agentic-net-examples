// Title: Copy a worksheet from one Excel workbook to another while preserving formulas using Aspose.Cells for .NET (C#)
// Common Searches: asp.net copy worksheet to another workbook preserving formulas with Aspose.Cells C# | how to duplicate a worksheet from source Excel file to target workbook using Aspose.Cells | C# example of copying first worksheet by name to existing workbook with formulas intact | Aspose.Cells AddCopy method usage for moving worksheets between workbooks
// Tags: Worksheet.AddCopy method Aspose.Cells | copy worksheet between workbooks C# | preserve formulas during worksheet copy | load existing workbook or create new Aspose.Cells | copy first worksheet by name C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads a source Excel file, opens or creates a target workbook, retrieves the first worksheet from the source, copies it to the target using the Worksheet.AddCopy method (which retains all formulas), and saves the target workbook, with error handling for missing files.
class WorksheetCopyExample
{
    static void Main()
    {
        try
        {
            // Paths for source and target workbooks
            string sourcePath = "SourceWorkbook.xlsx";
            string targetPath = "TargetWorkbook.xlsx";

            // Load the source workbook if it exists; otherwise throw an informative exception
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source workbook not found: {sourcePath}");

            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Load the target workbook if it exists; otherwise create a new empty workbook
            Workbook targetWorkbook = File.Exists(targetPath) ? new Workbook(targetPath) : new Workbook();

            // Get the worksheet to be copied from the source workbook (e.g., the first worksheet)
            Worksheet sourceWorksheet = sourceWorkbook.Worksheets[0];

            // Copy the worksheet to the target workbook using the worksheet name overload
            targetWorkbook.Worksheets.AddCopy(sourceWorksheet.Name);

            // Save the target workbook with the newly copied worksheet
            targetWorkbook.Save(targetPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
