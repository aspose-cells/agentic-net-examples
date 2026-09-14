// Title: Protect an Excel worksheet with no password using Aspose.Cells for .NET and verify its protection status
// AI Prompts: Generate C# code that calls Worksheet.Protect with ProtectionType.All and an empty password, checks Worksheet.IsProtected, and saves the workbook. | Demonstrate applying full sheet protection in Aspose.Cells without providing a password and confirming the IsProtected flag before saving.
// Common Searches: C# code to protect an Excel sheet with an empty password using Aspose.Cells | Worksheet.Protect empty string returns true IsProtected | How to use ProtectionType.All to lock all sheet features in Aspose.Cells | Saving an Excel file after applying sheet protection with no password using Aspose.Cells
// Tags: worksheet.Protect empty password Aspose.Cells | ProtectionType.All full sheet protection .NET | Worksheet.IsProtected verification Aspose.Cells | save workbook after sheet protection Aspose.Cells | Aspose.Cells worksheet protection without password

using System;
using Aspose.Cells;

// The program creates a new workbook, names the first worksheet, applies full protection with ProtectionType.All using an empty password, checks the IsProtected flag, and saves the file as ProtectedEmptyPassword.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (or add a new one)
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "ProtectedSheet";

            // Protect the worksheet without a password (empty password)
            // Use ProtectionType.All to protect all aspects of the sheet
            worksheet.Protect(ProtectionType.All);
            Console.WriteLine("Worksheet protected with empty password successfully.");

            // Verify the protection status
            bool isProtected = worksheet.IsProtected;
            Console.WriteLine("Is worksheet protected? " + isProtected);

            // Save the workbook
            string outputPath = "ProtectedEmptyPassword.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
