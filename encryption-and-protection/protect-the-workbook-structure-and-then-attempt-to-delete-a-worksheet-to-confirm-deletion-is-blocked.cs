// Title: Apply password‑protected structure protection to an Aspose.Cells workbook and confirm worksheet deletion is blocked in C#
// Common Searches: Aspose.Cells C# protect workbook structure with password and prevent sheet deletion | how to catch CellsException when trying to remove a worksheet from a protected workbook | verify worksheet count remains unchanged after structure protection in Aspose.Cells | C# code example for protecting workbook structure and testing delete restriction | Aspose.Cells workbook protection API usage for preventing worksheet removal
// Tags: enable structure protection on workbook Aspose.Cells | prevent worksheet removal with password protection | handle CellsException for protected workbook operations | Aspose.Cells workbook protection API example | verify worksheet count after structure protection | save protected workbook Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a new Workbook, adds an extra sheet, applies structure protection with a password, and then attempts to delete the first worksheet. Because the workbook structure is protected, the removal throws a CellsException, which is caught and reported. The worksheet count is displayed to confirm it remains unchanged, and the workbook is saved as 'ProtectedWorkbook.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add an extra worksheet for testing
            workbook.Worksheets.Add("Sheet2");

            // Protect the workbook structure with a password
            workbook.Protect(ProtectionType.Structure, "SecretPassword");

            // Attempt to delete the first worksheet
            try
            {
                workbook.Worksheets.RemoveAt(0);
                Console.WriteLine("Worksheet deleted (unexpected).");
            }
            catch (CellsException ex)
            {
                // Expected: deletion is blocked because the structure is protected
                Console.WriteLine("Deletion blocked: " + ex.Message);
            }

            // Verify that the worksheet count remains unchanged (should be 2)
            Console.WriteLine("Worksheet count: " + workbook.Worksheets.Count);

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
