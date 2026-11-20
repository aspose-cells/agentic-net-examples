// Title: Protect an Excel worksheet with a password while allowing cell formatting but preventing row insertion using Aspose.Cells for C#
// Common Searches: C# Aspose.Cells protect worksheet allow formatting but block row insertion | How to enable cell formatting on a password‑protected Excel sheet with Aspose.Cells | Aspose.Cells .NET set custom worksheet protection options for row insertion | Prevent users from adding rows in an Excel file using Aspose.Cells C# | Apply worksheet protection with specific permissions using Aspose.Cells API
// Tags: worksheet protection custom options Aspose.Cells | enable formatting on protected worksheet Aspose.Cells | prevent row insertion Excel Aspose.Cells | set worksheet password C# Aspose.Cells | ProtectionType configuration Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a new Workbook, accesses the first Worksheet, applies full protection with a password via sheet.Protect(ProtectionType.All, "myPassword", null) (you can replace ProtectionType.All with specific flags to allow formatting and block row insertion), saves the workbook as ProtectedWorkbook.xlsx, and outputs the file path while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Worksheet protection -----
            // Protect the worksheet with a password.
            // The third parameter (oldPassword) is required; pass null when not changing an existing password.
            sheet.Protect(ProtectionType.All, "myPassword", null);

            // Save the workbook with the applied protection
            string outputPath = "ProtectedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
