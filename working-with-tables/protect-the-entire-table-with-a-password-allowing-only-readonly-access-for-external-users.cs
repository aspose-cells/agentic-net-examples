// Title: Apply read‑only password protection to an entire worksheet table with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to protect all cells of a worksheet with a password that grants only view (read‑only) rights. | Show how to call Worksheet.Protect with ProtectionType.All and a custom password, then save the workbook as an .xlsx file. | Adapt an existing Aspose.Cells workbook to add read‑only protection for external users without altering the existing data.
// Common Searches: Aspose.Cells C# protect entire worksheet read only password example | How to set ProtectionType.All with a password in Aspose.Cells .NET | Make Excel sheet view‑only for external users using Aspose.Cells | C# code to apply password protection to a table in an Excel workbook with Aspose.Cells | Save a password‑protected workbook using Aspose.Cells for .NET
// Tags: worksheet password protection Aspose.Cells | read‑only worksheet protection C# | ProtectionType.All Aspose.Cells example | save password‑protected Excel file .NET | apply password to Excel table Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, optionally adds sample data, protects the entire worksheet with a password using Worksheet.Protect(ProtectionType.All,…), and saves the file as ProtectedTable.xlsx via Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet (the table to protect)
            var sheet = workbook.Worksheets[0];

            // Example data – can be omitted if the sheet already contains data
            sheet.Cells["A1"].PutValue("Sample Data");

            // Protect the worksheet with a password.
            // The third argument (oldPassword) is required; pass an empty string when not changing it.
            sheet.Protect(ProtectionType.All, "YourPasswordHere", string.Empty);

            // Save the protected workbook
            string outputPath = "ProtectedTable.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
