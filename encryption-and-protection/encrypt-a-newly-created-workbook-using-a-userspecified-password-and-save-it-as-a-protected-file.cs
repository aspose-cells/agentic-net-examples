// Title: Encrypt a newly created Excel workbook with a custom password using Aspose.Cells for .NET (C#)
// Common Searches: how to set a password on a new Excel workbook using Aspose.Cells in C# | C# code to encrypt an XLSX file with Aspose.Cells before saving | Aspose.Cells protect workbook with user‑defined password programmatically | save a password‑protected Excel file using Aspose.Cells .NET API | assign Workbook.Settings.Password for workbook encryption Aspose.Cells
// Tags: Aspose.Cells workbook password encryption | C# encrypt Excel file with Aspose.Cells | Workbook.Settings.Password property usage | save protected XLSX using Aspose.Cells | programmatic Excel file protection .NET

using System;
using Aspose.Cells;

// Creates a new Workbook, optionally adds data, assigns a user‑specified password via Workbook.Settings.Password to encrypt the file, and saves it as a protected XLSX using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // (Optional) Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample");
        sheet.Cells["B1"].PutValue("Data");

        // Set the password that will encrypt the file
        // This password is required to open the workbook later
        workbook.Settings.Password = "UserSpecifiedPassword";

        // Save the encrypted workbook to disk
        workbook.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
