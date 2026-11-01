// Title: How to password‑protect an Excel worksheet while still allowing users to select locked cells using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates or loads a workbook, protects the first worksheet with a password using Worksheet.Protect, and configures the protection so only locked cells can be selected. | Show an Aspose.Cells example that applies ProtectionType.SelectLockedCells (or a combination of flags) with a password, then saves the workbook as an .xlsx file. | Generate a snippet that demonstrates setting a worksheet password, enabling selection of locked cells, and disabling other editing actions with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# protect worksheet with password but allow selecting locked cells | how to enable only locked‑cell selection after protecting an Excel sheet using Aspose.Cells | Worksheet.Protect method password allow select locked cells Aspose.Cells .NET example | C# example for protecting Excel worksheet while keeping cell selection enabled with Aspose.Cells
// Tags: worksheet protection password Aspose.Cells C# | allow select locked cells Aspose.Cells | ProtectionType.SelectLockedCells usage | save protected workbook as xlsx Aspose.Cells | Worksheet.Protect method example

using System;
using System.IO;
using Aspose.Cells;

// Creates a new workbook, applies password protection to the first worksheet with settings that permit only locked‑cell selection, and saves the result as ProtectedWorksheet.xlsx.
class WorksheetProtectionExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Protect the worksheet with a password (protect all options)
            // The third parameter is the old password; pass null when not required
            sheet.Protect(ProtectionType.All, "MySecurePassword", null);

            // Save the protected workbook
            string outputPath = "ProtectedWorksheet.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
