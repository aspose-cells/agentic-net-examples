// Title: How to protect an Excel worksheet with a password while still allowing cell formatting using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# protect worksheet with password but keep format cells enabled | Set worksheet protection password and allow cell formatting in Aspose.Cells .NET | Enable format cells option when protecting an Excel sheet with Aspose.Cells | C# code to protect Excel worksheet and still let users change cell formatting using Aspose.Cells
// Tags: worksheet protection password Aspose.Cells | allow cell formatting on protected sheet Aspose.Cells | Aspose.Cells Protect method ProtectionType.All | C# Excel worksheet protection example | enable format cells option Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, renames the first worksheet, applies full protection with a password using sheet.Protect(ProtectionType.All, "MySecretPassword", string.Empty) while keeping the format‑cells permission enabled, saves the file as ProtectedWorkbook.xlsx, and handles any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (or add a new one if needed)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ProtectedSheet";

            // Protect the worksheet with a password.
            // The third parameter is the old password (empty because the sheet is not yet protected).
            sheet.Protect(ProtectionType.All, "MySecretPassword", string.Empty);

            // Save the workbook to a file
            string outputPath = "ProtectedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
