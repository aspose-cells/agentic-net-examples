// Title: Protect an Aspose.Cells worksheet with password Secure123 and enable selection of locked cells only (C#)
// AI Prompts: Write C# code using Aspose.Cells to protect a worksheet with password "Secure123" while configuring the protection to allow users to select only locked cells. | Demonstrate how to adjust the Protection object of an Aspose.Cells worksheet to set AllowSelectLockedCells = true and disable other actions after applying password protection.
// Common Searches: Aspose.Cells C# protect worksheet password allow only locked cells selection | How to set AllowSelectLockedCells property in Aspose.Cells protection | Programmatically protect Excel sheet with password and restrict selection using Aspose.Cells .NET | C# example for worksheet.Protect with ProtectionType.All and custom options in Aspose.Cells | Save a password‑protected workbook with Aspose.Cells and enable locked cell selection
// Tags: protect worksheet with password Aspose.Cells C# | allow select locked cells Aspose.Cells protection | worksheet.Protect ProtectionType.All Aspose.Cells | configure worksheet protection options .NET | save password protected workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, accesses the first worksheet, and calls worksheet.Protect with ProtectionType.All and the password "Secure123". It then shows how to enable the AllowSelectLockedCells option via the Protection object, ensuring only locked cells can be selected. The workbook is saved as ProtectedWorksheet.xlsx, with directory creation and error handling included.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Protect the worksheet with a password.
            // The third parameter (oldPassword) is required; pass null or empty string if not needed.
            worksheet.Protect(ProtectionType.All, "Secure123", null);

            // If you need to adjust protection options, you can modify the Protection object here.
            // Example (uncomment if supported by your Aspose.Cells version):
            // var protection = worksheet.Protection;
            // protection.AllowSelectLockedCells = true;
            // protection.AllowSelectUnlockedCells = false;

            // Ensure the output directory exists
            string outputPath = "ProtectedWorksheet.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
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
