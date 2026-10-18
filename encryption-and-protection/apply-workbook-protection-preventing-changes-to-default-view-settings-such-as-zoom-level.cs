// Title: Protect Excel workbook window settings such as zoom level using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells protect workbook window settings from user changes | Set password protection for Excel view options using Aspose.Cells .NET | Prevent zoom level modification in generated Excel file with Aspose.Cells | How to lock default view settings in an Excel workbook via Aspose.Cells C# | Apply workbook protection to window options in Aspose.Cells for .NET
// Tags: Aspose.Cells window settings lock | Excel view options protection C# | Workbook.Protect with ProtectionType.Windows | password protected workbook windows Aspose.Cells | lock zoom level in generated XLSX using Aspose.Cells

using System;
using Aspose.Cells;

// Demonstrates creating or loading an Excel workbook in C#, applying password‑protected protection to the workbook's window settings (including zoom level) with the ProtectionType.Windows flag, and saving the file.
class WorkbookProtectionExample
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        // Uncomment the line below to load an existing file instead of creating a new workbook
        // Workbook workbook = new Workbook("input.xlsx");
        Workbook workbook = new Workbook();

        // Apply protection to the workbook's window settings (e.g., zoom level, view options)
        // The ProtectionType.Windows flag prevents changes to default view settings
        // Optionally, provide a password to secure the protection
        workbook.Protect(ProtectionType.Windows, "MySecurePassword");

        // Save the protected workbook
        workbook.Save("ProtectedWorkbook.xlsx");
    }
}
