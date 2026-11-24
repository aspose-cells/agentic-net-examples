// Title: How to add an external named range that points to another workbook for cross‑file formulas with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# create named range that references cells in a different workbook | set RefersTo property for external workbook range using Aspose.Cells .NET | cross‑workbook named range example with Aspose.Cells in C# | define external reference string for a named range in Aspose.Cells | save workbook containing external named range using Aspose.Cells C#
// Tags: Aspose.Cells external named range C# | cross workbook reference using RefersTo | named range external workbook Aspose.Cells | define external reference string Aspose.Cells | save workbook with external named range

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, builds an external reference string that points to cells A1:B10 on Sheet1 of SourceWorkbook.xlsx, adds a named range called "ExternalData" to the target workbook, assigns the external reference to the range's RefersTo property, and saves the result as TargetWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook that will contain the named range
            Workbook targetWorkbook = new Workbook();

            // Define the external reference.
            // Format: '[WorkbookName]SheetName'!CellRange
            // Example references cells A1:B10 on Sheet1 of SourceWorkbook.xlsx
            string externalReference = "'[SourceWorkbook.xlsx]Sheet1'!$A$1:$B$10";

            // Add the named range to the target workbook.
            // First add a placeholder name, then set its RefersTo property to the external reference.
            int nameIndex = targetWorkbook.Worksheets.Names.Add("ExternalData");
            Name externalName = targetWorkbook.Worksheets.Names[nameIndex];
            externalName.RefersTo = externalReference;

            // Save the workbook that now holds the cross‑file named range
            string outputPath = "TargetWorkbook.xlsx";
            targetWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
