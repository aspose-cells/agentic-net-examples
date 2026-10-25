// Title: Protect an Excel worksheet with password Secure123 while allowing only selection of locked cells using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# protect worksheet with password and allow only select locked cells | How to set worksheet protection to select locked cells only using Aspose.Cells .NET | C# example for protecting an Excel sheet with password Secure123 in Aspose.Cells | Aspose.Cells protect sheet and restrict editing but enable selection of locked cells | Saving a password‑protected workbook as .xlsx with Aspose.Cells in C#
// Tags: Aspose.Cells Protect method with password C# | select locked cells only ProtectionType Aspose.Cells | save workbook as Xlsx after sheet protection .NET | C# worksheet protection using ProtectionType.All | Aspose.Cells example for sheet protection and cell selection

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, accesses the first Worksheet, applies full protection with the password "Secure123" using the Protect method, ensures the output directory exists, and saves the workbook as ProtectedWorksheet.xlsx while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Protect the worksheet with password "Secure123"
            // The third parameter is the old password (empty because the sheet is not previously protected)
            sheet.Protect(ProtectionType.All, "Secure123", string.Empty);

            // Save the workbook (save rule)
            string outputPath = "ProtectedWorksheet.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
