// Title: Unprotect an Excel worksheet, edit cell values, and re‑protect it with the original password using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to remove worksheet protection, change cells, and apply the same password | example code for unprotecting and re‑protecting an Excel sheet with ProtectionType.All in Aspose.Cells | modify cell A1 and B2 after unprotecting a protected worksheet using Aspose.Cells for .NET | Aspose.Cells protect worksheet with password after editing data | load workbook, unprotect first sheet, update values, save with protection Aspose.Cells C#
// Tags: Aspose.Cells unprotect worksheet C# | Aspose.Cells protect worksheet with password | ProtectionType.All usage Aspose.Cells | edit cells after worksheet unprotection Aspose.Cells | load workbook save with protection Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing Excel file, removes protection from the first worksheet, updates cells A1 and B2, then re‑applies full protection using the original password, and finally saves the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string password = "MySecretPassword";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Unprotect the worksheet (no password needed for removal)
            sheet.Unprotect();

            // ----- Modify cell values -----
            sheet.Cells["A1"].PutValue("New Text");
            sheet.Cells["B2"].PutValue(12345);

            // Re‑protect the worksheet with the same password and all protection types
            // The overload requires oldPassword; pass null when there is no previous password
            sheet.Protect(ProtectionType.All, password, null);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
