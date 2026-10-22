// Title: Protect an Excel worksheet with a password while still allowing cell selection using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells protect worksheet with password but keep cells selectable | How to enable cell selection on a password‑protected sheet using Aspose.Cells | Aspose.Cells set worksheet protection allow cell selection .NET | Protect Excel sheet with password in C# without disabling cell selection | Aspose.Cells worksheet protection options for selecting cells
// Tags: Aspose.Cells worksheet protection with password | enable cell selection on protected sheet C# | prevent editing of cells using Aspose.Cells Protect | ProtectionType.All usage Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first Worksheet, applies full protection with a password via sheet.Protect(ProtectionType.All, "myPassword", null) which blocks editing while still permitting cell selection, ensures the output directory exists, saves the file as ProtectedWorksheet.xlsx, and includes basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Protect the entire sheet with a password.
            // The third argument (oldPassword) is not required, so pass null.
            sheet.Protect(ProtectionType.All, "myPassword", null);

            // Determine output path and ensure the directory exists
            string outputPath = "ProtectedWorksheet.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
