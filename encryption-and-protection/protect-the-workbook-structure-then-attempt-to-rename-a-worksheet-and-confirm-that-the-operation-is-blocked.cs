// Title: How to protect workbook structure and prevent worksheet renaming with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# protect workbook structure with password and block sheet rename | prevent worksheet rename after protecting workbook structure using Aspose.Cells | C# code to lock Excel workbook structure and catch rename exception with Aspose.Cells | how to enable structure protection in Aspose.Cells and handle rename errors | Aspose.Cells protect workbook structure then attempt to rename sheet throws exception
// Tags: protect workbook structure Aspose.Cells | block worksheet rename Aspose.Cells | worksheet rename exception handling C# | password protected Excel workbook Aspose.Cells | structure protection API Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new Workbook, sets an initial worksheet name, applies password‑protected structure protection, then tries to rename the sheet. The rename operation throws an exception, which is caught and reported, demonstrating that structure protection blocks worksheet renaming. Finally, the protected workbook is saved.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the default worksheet and set an initial name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Original";

            // Protect the workbook structure with a password
            workbook.Protect(ProtectionType.Structure, "secret");

            // Attempt to rename the worksheet; this should be blocked
            try
            {
                sheet.Name = "Renamed";
                Console.WriteLine("Worksheet renamed successfully (unexpected).");
            }
            catch (Exception ex)
            {
                // Expected outcome: operation is blocked because the structure is protected
                Console.WriteLine("Rename operation blocked: " + ex.Message);
            }

            // Save the workbook (optional)
            workbook.Save("ProtectedWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            // General exception handling for unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
