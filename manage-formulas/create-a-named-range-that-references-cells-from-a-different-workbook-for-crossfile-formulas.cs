// Title: Define an external named range that points to another workbook and use it in a SUM formula with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to add a named range whose RefersTo points to a range in a separate Excel file, then place a SUM formula that references that named range. | Show how to build the external reference string for a Name object, assign it, and save the workbook that contains the external named range.
// Common Searches: asp.net cells create named range referencing another workbook | c# aspose.cells external named range formula example | how to set RefersTo property for external range in Aspose.Cells | using external workbook range in SUM formula with Aspose.Cells C#
// Tags: Aspose.Cells external named range definition | C# add RefersTo external reference | cross‑workbook range formula Aspose.Cells | named range external workbook Excel | SUM formula using external named range

using Aspose.Cells;
using System;
using System.IO;

// C# program that creates a new workbook, defines a named range pointing to cells A1:C10 in an external file (Data.xlsx) via the RefersTo property, inserts a SUM formula referencing that named range, and saves the workbook as MainWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook that will hold the external named range definition
            Workbook mainWb = new Workbook();

            // Define the external reference string.
            // Format: '[ExternalFile.xlsx]SheetName'!$StartCell:$EndCell
            // Example assumes the external workbook is "Data.xlsx" in the same directory.
            string externalRef = "'[Data.xlsx]Sheet1'!$A$1:$C$10";

            // Add a named range with the given name.
            // The RefersTo property is set separately to avoid overload issues.
            int nameIndex = mainWb.Worksheets.Names.Add("ExternalData");
            Name externalName = mainWb.Worksheets.Names[nameIndex];
            externalName.RefersTo = externalRef;

            Console.WriteLine($"Created named range: {externalName.Text} -> {externalName.RefersTo}");

            // Example usage: place a formula that references the external named range
            Worksheet ws = mainWb.Worksheets[0];
            ws.Cells["A1"].Formula = "=SUM(ExternalData)";

            // Save the workbook containing the external named range
            string outputPath = "MainWorkbook.xlsx";
            mainWb.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
