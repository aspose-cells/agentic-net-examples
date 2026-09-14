// Title: Save an XLSX workbook with worksheet protection and open password using Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, protect the first worksheet with a password, assign an open password to the workbook, and save it as an encrypted .xlsx file with Aspose.Cells in C#. | Apply full sheet protection and workbook-level encryption, then export the workbook to an .xlsx file using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set worksheet password and workbook open password before saving | how to encrypt an Excel file with both sheet and workbook passwords using Aspose.Cells .NET | save protected workbook as .xlsx with Aspose.Cells example in C#
// Tags: worksheet protection password Aspose.Cells C# | workbook open password encryption Aspose.Cells .NET | save encrypted XLSX Aspose.Cells C# | ProtectionType.All usage Aspose.Cells | Workbook.Settings.Password Aspose.Cells

using Aspose.Cells;
using System;

namespace AsposeCellsExample
{
    // The example creates a new workbook, writes "Hello" and "World" to cells A1 and B1, protects the first worksheet with a password using ProtectionType.All, sets an open password for the workbook, and saves the result as ProtectedEncrypted.xlsx in XLSX format.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet and add sample data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B1"].PutValue("World");

                // Protect the worksheet with a password (all protection types)
                // The third parameter (oldPassword) is not needed here, so pass null
                sheet.Protect(ProtectionType.All, "SheetPassword", null);

                // Apply workbook-level protection (open password)
                workbook.Settings.Password = "OpenPassword";

                // Save the workbook as an XLSX file
                workbook.Save("ProtectedEncrypted.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
