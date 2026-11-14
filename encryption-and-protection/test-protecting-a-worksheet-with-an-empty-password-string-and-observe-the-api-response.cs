// Title: Protect an Excel worksheet with an empty password using Aspose.Cells for .NET and verify the IsProtected flag
// Common Searches: Aspose.Cells protect worksheet with empty password C# example | How does sheet.Protect behave when password is omitted in Aspose.Cells | Check IsProtected property after calling Protect without password in Aspose.Cells | Saving workbook after protecting worksheet without password using Aspose.Cells | C# code to protect Excel sheet without password using Aspose.Cells
// Tags: worksheet protection without password Aspose.Cells | sheet.Protect method usage Aspose.Cells | IsProtected property check Aspose.Cells | saving protected workbook Aspose.Cells | ProtectionType.All example Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a new Workbook, renames the first worksheet, applies sheet.Protect(ProtectionType.All) without providing a password, prints the worksheet's IsProtected status, saves the file as ProtectedEmptyPassword.xlsx, and catches any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ProtectedSheet";

            // Protect the worksheet without a password
            sheet.Protect(ProtectionType.All);

            // Display protection status
            Console.WriteLine($"Worksheet '{sheet.Name}' protected: {sheet.IsProtected}");

            // Save the workbook
            string outputPath = "ProtectedEmptyPassword.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
