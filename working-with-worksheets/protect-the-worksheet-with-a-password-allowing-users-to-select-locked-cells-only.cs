// Title: How to password‑protect an Excel worksheet but still allow selection of locked cells with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to protect a worksheet with a password while keeping the ability to select locked cells. | Show the required parameters for the Aspose.Cells Protect method to enable full protection yet allow cell selection in an Excel file.
// Common Searches: Aspose.Cells C# protect worksheet password allow selecting locked cells | set worksheet protection type all but keep cell selection enabled Aspose.Cells | how to enable cell selection on a password‑protected sheet using Aspose.Cells .NET | C# Aspose.Cells protect sheet without disabling cell selection | Excel worksheet protection with password while still allowing user to click locked cells Aspose
// Tags: Aspose.Cells worksheet password protection API | preserve cell selection on protected sheet | ProtectionType.All usage in C# Aspose.Cells | save password‑protected Excel workbook .NET | configure worksheet protection settings Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, renames the first worksheet, applies full protection with a password using sheet.Protect(ProtectionType.All, "MySecretPassword", null), ensures the output directory exists, and saves the file as ProtectedWorkbook.xlsx.
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

            // Protect the worksheet with a password (protect all aspects)
            // The third parameter is the old password; null is used when there is no existing password.
            sheet.Protect(ProtectionType.All, "MySecretPassword", null);

            // Define output path and ensure its directory exists
            string outputPath = "ProtectedWorkbook.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
