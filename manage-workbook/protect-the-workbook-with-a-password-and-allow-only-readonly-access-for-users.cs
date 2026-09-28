// Title: How to password‑protect an Excel workbook for read‑only access using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to apply a password to a workbook's structure and windows, then save it as a read‑only .xlsx file. | Show an example of calling workbook.Protect with ProtectionType.All and a password, followed by saving the protected workbook in C#. | Demonstrate creating a workbook, inserting data, setting a password for read‑only access, and exporting the file with Aspose.Cells.
// Common Searches: Aspose.Cells C# protect workbook with password and restrict editing | Set read‑only mode for an Excel file using Aspose.Cells .NET | How to apply structure and windows protection to an .xlsx with Aspose.Cells | C# example for password‑protected Excel workbook using Aspose.Cells
// Tags: Aspose.Cells workbook.Protect method | C# Excel workbook password protection | read‑only .xlsx Aspose.Cells | protect workbook structure windows Aspose | save protected workbook .NET

using Aspose.Cells;
using System;

// The sample creates a new Workbook, adds data to the first worksheet, protects the workbook's structure and windows with a password using workbook.Protect(ProtectionType.All, password), and saves the resulting read‑only Excel file as 'ProtectedWorkbook.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["A2"].PutValue("Data");

            // Protect the workbook (structure and windows) with a password
            string password = "MySecretPassword";
            workbook.Protect(ProtectionType.All, password);

            // Save the protected workbook
            string outputPath = "ProtectedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
