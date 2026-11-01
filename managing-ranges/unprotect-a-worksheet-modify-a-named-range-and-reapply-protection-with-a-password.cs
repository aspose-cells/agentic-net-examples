// Title: Unprotect an Excel worksheet, modify a named range, and re‑apply password protection using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells unprotect worksheet, change named range value, then protect with password | How to edit a named range after removing sheet protection with Aspose.Cells for .NET | Aspose.Cells example: unprotect sheet, update MyRange cell, protect sheet with password | Programmatically modify a named range in a protected Excel file using Aspose.Cells C# | Re‑apply full protection to a worksheet after updating a named range with Aspose.Cells
// Tags: worksheet unprotect Aspose.Cells C# | named range edit Aspose.Cells | worksheet protect password Aspose.Cells | Aspose.Cells GetRange modify cell | save workbook Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample loads an Excel workbook, removes protection from the first worksheet, retrieves the named range "MyRange", updates its first cell, reapplies full protection with the password "MyPassword", and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Unprotect the worksheet (if it was protected)
            sheet.Unprotect();

            // Retrieve the named range from the workbook's name collection
            Name namedRange = workbook.Worksheets.Names["MyRange"];
            if (namedRange != null)
            {
                // Get the range that the named range refers to
                AsposeRange range = namedRange.GetRange();
                if (range != null && range.RowCount > 0 && range.ColumnCount > 0)
                {
                    // Modify the first cell of the range
                    Cell cell = range[0, 0];
                    cell.PutValue("New Value");
                }
            }
            else
            {
                Console.WriteLine("Named range 'MyRange' not found.");
            }

            // Re‑apply protection with a password
            sheet.Protect(ProtectionType.All, "MyPassword", null);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
