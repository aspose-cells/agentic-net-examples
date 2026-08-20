// Title: Protect Workbook Structure and Block Sheet Renaming with Aspose.Cells for .NET
// Description: Creates a new Workbook, applies structure protection with a password, attempts to rename the first worksheet (throws an exception), verifies the sheet name remains unchanged, and saves the protected file as an XLSX document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells protect workbook structure | C# protect Excel workbook structure | prevent worksheet rename Aspose.Cells | structure protection .NET | verify sheet rename blocked | save protected workbook Aspose.Cells | Excel workbook password protection C#
// Common Searches: Aspose.Cells protect workbook structure C# | How to block sheet rename in Aspose.Cells | C# example to protect Excel workbook structure with password | Verify worksheet name unchanged after protection Aspose.Cells | Save password‑protected workbook using Aspose.Cells
// Developer Intent: Demonstrate applying password‑protected structure protection to an Excel workbook, attempting a sheet rename, catching the expected exception, confirming the name stays the same, and saving the file.
// Use Cases: Distribute a template where sheet names and order must stay fixed | Automated unit test to ensure structure protection blocks prohibited actions | Create a read‑only layout Excel file while allowing data entry | Compliance‑driven Excel files that prevent users from altering worksheet organization
// AI Prompts: Write C# code using Aspose.Cells to protect a workbook's structure with a password, try to rename the first worksheet, handle the exception, confirm the name didn't change, and save the file. | Explain step‑by‑step how to verify that a worksheet rename is blocked after applying structure protection in Aspose.Cells for .NET. | Provide a concise tutorial for protecting an Excel workbook's structure and testing the protection by attempting a sheet rename.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new Workbook, applies structure protection with a password, attempts to rename the first worksheet (throws an exception), verifies the sheet name remains unchanged, and saves the protected file as an XLSX document using Aspose.Cells for .NET.
    public class ProtectStructureAndAttemptRename
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Protect the workbook structure with a password
                workbook.Protect(ProtectionType.Structure, "myPassword");

                // Keep the original worksheet name for verification
                Worksheet sheet = workbook.Worksheets[0];
                string originalName = sheet.Name;

                // Attempt to rename the worksheet; this should be blocked because the structure is protected
                try
                {
                    sheet.Name = "RenamedSheet";
                    Console.WriteLine("Rename operation succeeded (unexpected).");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Rename operation blocked as expected: " + ex.Message);
                }

                // Verify that the worksheet name has not changed
                bool nameUnchanged = sheet.Name == originalName;
                Console.WriteLine($"Worksheet name unchanged: {nameUnchanged}");

                // Save the workbook (lifecycle rule: save)
                workbook.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved as ProtectedWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ProtectStructureAndAttemptRename.Run();
        }
    }
}
