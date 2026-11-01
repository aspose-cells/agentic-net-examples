// Title: How to protect an Excel worksheet with a password while still allowing cell formatting using Aspose.Cells for .NET and verify that locked cells cannot be edited programmatically
// Common Searches: Aspose.Cells protect worksheet password allow formatting C# example | check if a cell is locked after worksheet protection Aspose.Cells .NET | prevent programmatic changes to locked cells in Excel with Aspose.Cells | use ProtectionType.All to enable full worksheet protection in C# | save protected Excel file and test cell edit restriction using Aspose.Cells
// Tags: worksheet password protection with formatting Aspose.Cells | verify locked cell edit restriction Aspose.Cells | ProtectionType.All usage C# | cell style IsLocked property Aspose.Cells | save protected workbook Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, writes a value to cell A1, protects the first worksheet with a password using ProtectionType.All (enabling all protection options while still allowing formatting), saves the file, attempts to modify the locked cell programmatically to show that the API can bypass UI protection, and finally reads the cell's IsLocked style flag to confirm the lock status.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ProtectedSheet";

            // Add initial data to a locked cell (cells are locked by default)
            sheet.Cells["A1"].PutValue("Original Value");

            // Protect the worksheet with a password (all protection options enabled)
            // The overload requires the old password; pass null for a new protection
            sheet.Protect(ProtectionType.All, "SecretPassword", null);

            // Save the workbook (ensure the directory exists)
            string outputPath = "ProtectedSheet.xlsx";
            try
            {
                // Ensure the directory for the output file exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
                return;
            }

            // Attempt to edit the locked cell programmatically
            try
            {
                sheet.Cells["A1"].PutValue("Attempted Edit");
                Console.WriteLine("Programmatic edit succeeded (API bypasses UI protection).");
            }
            catch (CellsException ex)
            {
                Console.WriteLine("Edit blocked: " + ex.Message);
            }

            // Verify cell lock status
            Console.WriteLine($"Cell A1 locked: {sheet.Cells["A1"].GetStyle().IsLocked}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
