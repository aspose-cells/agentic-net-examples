// Title: Generate an XLSX file, merge cells R2:S2, and apply worksheet protection with a password using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new workbook with Aspose.Cells, merges the range R2:S2 on the first worksheet, protects the sheet with a password, and saves the result as an XLSX file. | Show how to enable worksheet protection so that the merged cells R2:S2 become read‑only, then export the workbook to XLSX using Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# merge cells R2:S2 and protect worksheet with password | how to create an XLSX file with a protected merged range using Aspose.Cells for .NET | C# Aspose.Cells protect merged cells read only while saving as XLSX | save workbook with merged cells and worksheet protection in Aspose.Cells
// Tags: Aspose.Cells merge specific range C# | apply worksheet protection Aspose.Cells | export workbook to XLSX format Aspose.Cells | read‑only merged cells Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Creates a new workbook, merges cells R2:S2, protects the worksheet with a password, and saves the file as MergedProtected.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Merge cells R2:S2 using a Range object
            var range = sheet.Cells.CreateRange("R2", "S2");
            range.Merge();

            // Protect the worksheet with a password (oldPassword is not required, pass empty string)
            sheet.Protect(ProtectionType.All, "password", string.Empty);

            // Define output file path
            string outputPath = "MergedProtected.xlsx";

            // Save the workbook as XLSX
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
