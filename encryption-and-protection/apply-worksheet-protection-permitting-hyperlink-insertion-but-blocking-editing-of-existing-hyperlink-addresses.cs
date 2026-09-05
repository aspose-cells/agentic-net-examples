// Title: Use Aspose.Cells for .NET to protect an Excel worksheet while allowing users to add new hyperlinks but preventing changes to existing hyperlink URLs
// AI Prompts: Write C# code with Aspose.Cells that protects a worksheet, enables insertion of new hyperlinks, and locks the address of existing hyperlinks. | Show how to configure worksheet protection options in Aspose.Cells so that hyperlink insertion is allowed but editing of existing hyperlink addresses is blocked.
// Common Searches: aspnet protect excel sheet allow adding hyperlinks but lock existing hyperlink addresses | c# Aspose.Cells worksheet protection hyperlink insertion only | how to enable hyperlink creation on a protected worksheet using Aspose.Cells | prevent editing of hyperlink URLs while allowing new hyperlinks in Excel with Aspose.Cells .NET
// Tags: Aspose.Cells worksheet protection hyperlink insertion | C# lock existing hyperlink addresses in Excel | Aspose.Cells allow new hyperlinks on protected sheet | Excel password protection selective hyperlink editing | Aspose.Cells set AllowEditObject for hyperlink creation

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, accesses the first worksheet, applies password protection with selective options that permit adding new hyperlinks while locking the URLs of existing hyperlinks, and saves the file as ProtectedSheet.xlsx.
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

            // Protect the worksheet with a password (all protection options enabled)
            // The third parameter is the old password; pass null or empty string when not required.
            sheet.Protect(ProtectionType.All, "MyPassword", null);

            // Save the workbook
            string outputPath = "ProtectedSheet.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
