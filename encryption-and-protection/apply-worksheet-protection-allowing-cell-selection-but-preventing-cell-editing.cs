// Title: Protect an Excel worksheet with Aspose.Cells for .NET while permitting selection of locked and unlocked cells
// Common Searches: Aspose.Cells C# protect worksheet but still allow selecting locked cells | How to enable cell selection on a protected Excel sheet using Aspose.Cells | Set AllowSelectingLockedCell and AllowSelectingUnlockedCell in Aspose.Cells .NET | Save a password‑protected workbook with selectable cells using Aspose.Cells | C# example for worksheet protection with selectable cells in Excel
// Tags: Aspose.Cells worksheet protection allow selecting cells | C# protect Excel sheet with password using ProtectionType.All | Aspose.Cells set AllowSelectingLockedCell property | Excel workbook save with selectable locked cells Aspose.Cells | Aspose.Cells Protect method configuration .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, applies password‑protected worksheet protection using ProtectionType.All, enables selection of both locked and unlocked cells via the AllowSelectingLockedCell and AllowSelectingUnlockedCell flags, and saves the file as ProtectedWorksheet.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Protect the worksheet with a password (prevents editing)
            // ProtectionType.All protects the entire sheet.
            // The third parameter is the old password; pass null when not required.
            sheet.Protect(ProtectionType.All, "SecurePassword", null);

            // Allow users to select both locked and unlocked cells while editing remains disabled
            sheet.Protection.AllowSelectingLockedCell = true;
            sheet.Protection.AllowSelectingUnlockedCell = true;

            // Define output file path
            string outputPath = "ProtectedWorksheet.xlsx";

            // Ensure the directory exists (if a directory part is present)
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook (lifecycle: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
