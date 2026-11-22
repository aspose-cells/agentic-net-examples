// Title: How to protect an Excel workbook’s structure with a custom password using Aspose.Cells for .NET and save it as XLSX
// Common Searches: C# code to lock Excel workbook structure with a password using Aspose.Cells | Aspose.Cells protect workbook structure and save as .xlsx in .NET | Set custom password for workbook protection with Aspose.Cells API | Example of workbook.Protect(ProtectionType.Structure) in C# | Save password‑protected Excel file using Aspose.Cells for .NET
// Tags: Aspose.Cells workbook structure protection C# | password-protected workbook structure Aspose.Cells | save protected Excel file as Xlsx Aspose | ProtectionType.Structure usage Aspose.Cells | Excel workbook lock structure example

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, adds data to cell A1, applies structure protection with a custom password via workbook.Protect(ProtectionType.Structure, password), and saves the password‑protected file as ProtectedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Optional: add some data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");

            // Protect the workbook structure with a custom password
            string password = "MySecretPassword";
            workbook.Protect(ProtectionType.Structure, password);

            // Determine output file path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ProtectedWorkbook.xlsx");

            // Save the protected workbook to a file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
