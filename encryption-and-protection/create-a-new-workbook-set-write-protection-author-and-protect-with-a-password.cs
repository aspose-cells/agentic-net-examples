// Title: Create an Excel workbook, set a write‑protection author, and apply password protection with Aspose.Cells for .NET (C#)
// Common Searches: c# aspocells set write protection author for new workbook | protect an Excel file with password using Aspose.Cells in .NET | how to assign write‑protection author and password to a workbook with Aspose.Cells | aspocells workbook write protection example in C# | save a password‑protected Excel workbook using Aspose.Cells for .NET
// Tags: aspocells set write protection author | aspocells password protect workbook | c# create protected excel workbook | aspocells workbook write protection settings | excel workbook password protection c#

using System;
using Aspose.Cells;

// The sample creates a new Workbook object, assigns "John Doe" as the write‑protection author, sets a password for write protection, and saves the workbook as a password‑protected Excel file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set write protection author
        workbook.Settings.WriteProtection.Author = "John Doe";

        // Protect the workbook with a password
        workbook.Settings.WriteProtection.Password = "MySecurePassword";

        // Save the workbook to a file
        workbook.Save("ProtectedWorkbook.xlsx");
    }
}
