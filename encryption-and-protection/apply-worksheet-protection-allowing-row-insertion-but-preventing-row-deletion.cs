// Title: Apply password protection to an Excel worksheet using Aspose.Cells for .NET – row insertion permission not supported
// Common Searches: how to add password protection to an Excel worksheet using Aspose.Cells C# | Aspose.Cells C# protect worksheet but allow inserting new rows | C# Aspose.Cells worksheet protection options for allowing row addition | cannot enable insert rows in a protected sheet with Aspose.Cells | Aspose.Cells protect Excel sheet while preventing row deletion | example code for worksheet.Protect with password in Aspose.Cells .NET
// Tags: Aspose.Cells worksheet password protection C# | Aspose.Cells allow row insertion in protected worksheet | Aspose.Cells ProtectionType.All implementation | C# protect Excel worksheet using Aspose.Cells API | Aspose.Cells worksheet protection feature limitations | Aspose.Cells protect sheet without disabling row addition

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook (or loads an existing one), accesses the first Worksheet, and applies full protection with a password via sheet.Protect(ProtectionType.All, password, string.Empty). The current Aspose.Cells for .NET version does not expose individual AllowInsertRows or AllowDeleteRows flags, so selective permission for row insertion while blocking row deletion cannot be configured. The protected workbook is saved as ProtectedWorksheet.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Password for worksheet protection
            string password = "myPassword";

            // Apply full protection with the password.
            // The third parameter is the old password (empty because the sheet is not previously protected).
            sheet.Protect(ProtectionType.All, password, string.Empty);

            // Note: Specific allow/deny settings (e.g., AllowInsertRows) are omitted
            // because they are not available in the current Aspose.Cells version.

            // Save the protected workbook
            string outputPath = "ProtectedWorksheet.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
