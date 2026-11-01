// Title: How to protect workbook structure and prevent worksheet deletion with Aspose.Cells for .NET
// Common Searches: Aspose.Cells protect workbook structure C# example | prevent deleting a sheet from a protected workbook using Aspose.Cells | catch CellsException when removing a worksheet from a protected workbook in .NET | protect Excel file structure without password with Aspose.Cells API | remove worksheet from a protected workbook throws exception Aspose.Cells
// Tags: workbook structure protection Aspose.Cells | prevent worksheet deletion C# | CellsException handling sheet removal | protect Excel workbook without password | Aspose.Cells blocked sheet removal | save protected workbook Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a Workbook, adds a second sheet, applies structure protection with an empty password, attempts to delete the first worksheet (which raises a CellsException), catches the exception to confirm the deletion is blocked, and saves the protected file as "ProtectedWorkbook.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a second worksheet for demonstration
            workbook.Worksheets.Add("Sheet2");

            // Protect the workbook structure (empty password)
            workbook.Protect(ProtectionType.Structure, string.Empty);

            // Attempt to delete the first worksheet
            try
            {
                workbook.Worksheets.RemoveAt(0);
                Console.WriteLine("Worksheet deleted (unexpected).");
            }
            catch (CellsException ex)
            {
                Console.WriteLine("Deletion blocked as expected: " + ex.Message);
            }

            // Save the workbook
            workbook.Save("ProtectedWorkbook.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
