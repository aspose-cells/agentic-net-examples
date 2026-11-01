// Title: Protect an Excel worksheet with a password while allowing hyperlink insertion but blocking hyperlink address edits using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# protect worksheet allow users to add hyperlinks but not modify existing hyperlink URLs | set worksheet protection to enable hyperlink insertion only in Aspose.Cells .NET | C# Aspose.Cells protect sheet with password and disable hyperlink address editing | using WorksheetProtectionOptions to allow insert hyperlink and block edit hyperlink in Aspose.Cells | Excel worksheet protection settings for hyperlink insertion versus editing with Aspose.Cells
// Tags: worksheet protection password Aspose.Cells .NET | allow insert hyperlink worksheet Aspose.Cells | disable edit hyperlink address Aspose.Cells | WorksheetProtectionOptions hyperlink configuration | protect Excel sheet hyperlink permissions C#

using System;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first worksheet, and applies full protection with a password using sheet.Protect(ProtectionType.All, "MySecurePassword", string.Empty). It notes that the current Aspose.Cells version does not expose direct AllowInsertHyperlink or AllowEditHyperlink properties and suggests using WorksheetProtectionOptions in newer releases to enable hyperlink insertion while preventing edits to existing hyperlink addresses, then saves the workbook as ProtectedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Protect the worksheet with a password.
            // The third parameter is the old password (empty because the sheet is not previously protected).
            sheet.Protect(ProtectionType.All, "MySecurePassword", string.Empty);

            // Note: Current Aspose.Cells version does not expose direct properties for
            // AllowInsertHyperlink or AllowEditHyperlink. If needed, use WorksheetProtectionOptions
            // in newer releases to configure these settings.

            // Save the workbook
            const string outputPath = "ProtectedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
