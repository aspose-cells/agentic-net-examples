// Title: Password‑protect an Excel workbook and enforce read‑only access using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set password to open Excel workbook and make it read‑only | How to protect workbook structure and windows with a password in Aspose.Cells | Aspose.Cells .NET example for workbook.Settings.Password and Protect method | Save an Excel file as read‑only with password protection using Aspose.Cells | Apply full protection to an .xlsx file programmatically with Aspose.Cells C#
// Tags: Aspose.Cells workbook password protection C# | Workbook.Protect API for Excel structure windows | set workbook opening password Aspose.Cells .NET | read‑only Excel file generation Aspose.Cells | protect Excel workbook with password using C#

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookProtectionDemo
{
    // The example creates or loads a Workbook, sets an opening password via workbook.Settings.Password, applies full protection (structure and windows) with workbook.Protect(ProtectionType.All, password), and saves the file as a read‑only, password‑protected .xlsx.
    class WorkbookProtectionExample
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Set a password required to open the workbook
                workbook.Settings.Password = "SecurePassword123";

                // Protect the workbook structure and windows with the same password
                workbook.Protect(ProtectionType.All, "SecurePassword123");

                // Save the protected workbook
                string outputPath = "ProtectedWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
