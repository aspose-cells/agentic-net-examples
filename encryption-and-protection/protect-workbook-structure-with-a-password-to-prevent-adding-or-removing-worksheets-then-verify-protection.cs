// Title: C# example: protect Excel workbook structure with a password using Aspose.Cells and confirm protection by attempting worksheet modifications
// Common Searches: Aspose.Cells C# protect workbook structure with password and prevent sheet addition | How to verify workbook structure protection after saving with Aspose.Cells .NET | C# code to block adding or deleting worksheets in an Excel file using Aspose.Cells | Aspose.Cells Protect(ProtectionType.Structure) usage example in .NET | Exception handling when trying to modify a password‑protected workbook with Aspose.Cells
// Tags: Aspose.Cells workbook structure lock C# | Password‑based Excel workbook protection Aspose.Cells | Prevent sheet add delete with Aspose.Cells .NET | Protect method with ProtectionType.Structure example | Structure protection validation using exceptions Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a Workbook, adds three sheets, applies structure protection with a password via wb.Protect(ProtectionType.Structure, password), saves the file, reloads it, and then attempts to add and remove worksheets; caught exceptions confirm that the protection successfully blocks those operations.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some worksheets
            Workbook wb = new Workbook();
            wb.Worksheets.Add("Sheet1");
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets.Add("Sheet3");

            // Protect the workbook structure with a password
            string password = "MySecretPwd";
            wb.Protect(ProtectionType.Structure, password);

            // Save the protected workbook
            string filePath = "ProtectedWorkbook.xlsx";
            wb.Save(filePath);
            Console.WriteLine($"Workbook saved to '{filePath}' with structure protection.");

            // Ensure the file exists before loading
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File '{filePath}' was not found.");
                return;
            }

            // Load the workbook to verify protection
            Workbook loadedWb = new Workbook(filePath);

            // Verify that structure protection is enabled (if supported)
            // The Protection property may not be available in older versions;
            // therefore we rely on the exception handling when modifying the workbook.
            Console.WriteLine("Structure protection verification: attempted operations will confirm it.");

            // Attempt to add a new worksheet (should fail)
            try
            {
                loadedWb.Worksheets.Add("NewSheet");
                Console.WriteLine("Worksheet added (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Cannot add worksheet: {ex.Message}");
            }

            // Attempt to remove an existing worksheet (should fail)
            try
            {
                loadedWb.Worksheets.RemoveAt(0);
                Console.WriteLine("Worksheet removed (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Cannot remove worksheet: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
