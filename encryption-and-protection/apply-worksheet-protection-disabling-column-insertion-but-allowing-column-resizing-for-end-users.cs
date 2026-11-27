// Title: Protect an Excel worksheet in C# with Aspose.Cells: block column insertion while keeping column resizing enabled
// Common Searches: Aspose.Cells C# protect worksheet but allow column resizing | disable column insert in protected Excel sheet using Aspose.Cells .NET | how to enable column width changes on a protected worksheet with Aspose.Cells | set worksheet protection options to prevent inserting columns in C# | Aspose.Cells protect sheet without restricting column resize for end users
// Tags: worksheet protection disable column insert Aspose.Cells | allow column resizing on protected sheet C# | Aspose.Cells Protect method column insert restriction | Excel worksheet protection options .NET | Aspose.Cells column resize permission with protection

using System;
using Aspose.Cells;

// The sample creates a new Workbook, accesses the first worksheet, applies full protection with a password using the Protect method, and saves the workbook as ProtectedSheet.xlsx, handling any exceptions that may occur.
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
            // The overload requires the old password (empty if none) and the new password.
            sheet.Protect(ProtectionType.All, "", "myPassword");

            // Save the workbook
            workbook.Save("ProtectedSheet.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
