// Title: Aspose.Cells .NET – Protect rows 20‑25 with a password using the EntireRow property
// Description: A C# sample that creates a workbook, selects rows 20‑25 via the EntireRow property (or an AllowEditRange), assigns a password, applies worksheet protection, and saves the result as ProtectedRows.xlsx.
// Keywords: Aspose.Cells EntireRow | protect specific rows C# | password protect rows Aspose.Cells | AllowEditRanges .NET | Excel row protection | C# Excel security | Aspose.Cells worksheet protection
// Common Searches: Aspose.Cells protect rows with password | C# protect rows 20 to 25 Excel | How to use EntireRow property Aspose.Cells | Set password for row range Aspose.Cells .NET | Excel row lock example Aspose.Cells
// Developer Intent: Secure rows 20‑25 of a worksheet so they can only be edited after entering a password.
// Use Cases: Prevent accidental changes to header or total rows in financial reports. | Restrict access to confidential data rows before sharing a spreadsheet with clients. | Lock template sections while allowing users to fill in other parts of the sheet.
// AI Prompts: Generate C# code that uses Aspose.Cells' EntireRow property to password‑protect rows 20‑25 while keeping other cells editable. | Show how to combine AllowEditRanges with worksheet protection to lock a specific row range in Aspose.Cells for .NET. | Provide an example of protecting multiple non‑contiguous row blocks, each with its own password, using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsProtectedRowsDemo
{
    // A C# sample that creates a workbook, selects rows 20‑25 via the EntireRow property (or an AllowEditRange), assigns a password, applies worksheet protection, and saves the result as ProtectedRows.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a protected range that covers rows 20 through 25 (zero‑based indices 19‑24)
                // The range starts at column A (index 0) and ends at column A (index 0)
                int rangeIndex = worksheet.AllowEditRanges.Add("Rows20to25", 19, 0, 24, 0);
                ProtectedRange protectedRange = worksheet.AllowEditRanges[rangeIndex];

                // Set a password for the protected range
                protectedRange.Password = "MySecretPassword";

                // Protect the entire worksheet (all protection types)
                worksheet.Protect(ProtectionType.All);

                // Save the workbook
                workbook.Save("ProtectedRows.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
