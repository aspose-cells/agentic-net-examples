using System;
using Aspose.Cells;

class TestEmptyPasswordProtection
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Protect the worksheet using an empty password string
        // The third parameter (oldPassword) is null because the sheet is not previously protected
        sheet.Protect(ProtectionType.All, "", null);

        // Observe protection-related properties
        Console.WriteLine("Worksheet.IsProtected: " + sheet.IsProtected);
        Console.WriteLine("Worksheet.Protection.IsProtectedWithPassword: " + sheet.Protection.IsProtectedWithPassword);
        Console.WriteLine("Worksheet.Protection.Password (should be empty): '" + sheet.Protection.Password + "'");

        // Attempt to unprotect without providing a password
        // According to the API, an empty password allows unprotecting without a password
        try
        {
            sheet.Unprotect(); // Should succeed
            Console.WriteLine("Unprotected successfully without password.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to unprotect without password: " + ex.Message);
        }

        // Verify that the worksheet is no longer protected
        Console.WriteLine("Worksheet.IsProtected after unprotect: " + sheet.IsProtected);

        // Save the workbook (required lifecycle step)
        workbook.Save("EmptyPasswordProtectedWorksheet.xlsx");
    }
}