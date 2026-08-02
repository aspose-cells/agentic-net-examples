// Title: Protect Excel Workbook Structure with a Password using Aspose.Cells for .NET
// Description: C# example that creates a new Workbook, applies structure‑level protection with a custom password via workbook.Protect(ProtectionType.Structure, "MySecretPassword"), saves the file as ProtectedWorkbook.xlsx (XLSX format), and releases resources.
// Keywords: Aspose.Cells protect workbook structure | C# workbook password protection | Excel structure protection .NET | Save password‑protected Excel file | ProtectionType.Structure Aspose | Aspose.Cells encryption and protection
// Common Searches: How to lock workbook structure with a password in Aspose.Cells C# | Save an Excel file with structure protection using Aspose.Cells | Aspose.Cells protect only workbook layout not cells | C# code to apply password to workbook structure
// Developer Intent: Apply a password to the workbook's structure and save the protected Excel file using Aspose.Cells for .NET.
// Use Cases: Distribute a template where users can edit cells but cannot add, delete, or rename sheets. | Secure the layout of a financial model before publishing to a shared repository. | Prevent accidental sheet modifications in automated report generation pipelines.
// AI Prompts: Generate C# code that protects both workbook structure and windows with a custom password using Aspose.Cells and saves as .xlsb. | Show how to detect existing workbook protection before applying structure protection in Aspose.Cells for .NET. | Explain how to combine workbook structure protection with file‑level encryption in Aspose.Cells.

using System;
using Aspose.Cells;

// C# example that creates a new Workbook, applies structure‑level protection with a custom password via workbook.Protect(ProtectionType.Structure, "MySecretPassword"), saves the file as ProtectedWorkbook.xlsx (XLSX format), and releases resources.
class ProtectWorkbookStructure
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Protect the workbook structure with a custom password
        workbook.Protect(ProtectionType.Structure, "MySecretPassword");

        // Save the protected workbook to a file
        workbook.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);

        // Release resources
        workbook.Dispose();
    }
}
