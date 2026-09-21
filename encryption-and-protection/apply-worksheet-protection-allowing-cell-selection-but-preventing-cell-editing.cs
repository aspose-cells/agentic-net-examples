// Title: How to protect an Excel worksheet with Aspose.Cells for .NET so users can select cells but cannot edit them
// AI Prompts: Generate C# using Aspose.Cells that locks a sheet for editing while keeping both locked and unlocked cells selectable, without requiring a password. | Outline the steps to configure worksheet protection options in Aspose.Cells so cell selection remains enabled but all editing actions are blocked, then save the workbook.
// Common Searches: Aspose.Cells C# protect worksheet allow cell selection but prevent editing | How to enable selection of locked cells in an Excel file using Aspose.Cells .NET | Save Excel workbook with protection that blocks edits yet lets users select cells in C#
// Tags: Aspose.Cells worksheet.Protect selection options | C# disable editing on Excel sheet Aspose.Cells | Excel protection without password using Aspose.Cells | Set ProtectionType.SelectLockedCells in Aspose.Cells | Aspose.Cells configure worksheet protection C#

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, retrieves the first worksheet, applies full protection via worksheet.Protect(ProtectionType.All) which permits selection of both locked and unlocked cells while preventing any modifications, ensures the output directory exists, and saves the workbook as ProtectedWorksheet.xlsx, with exception handling for robustness.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            // Protect the worksheet without a password (protect all options)
            worksheet.Protect(ProtectionType.All);

            // Define output path
            string outputPath = "ProtectedWorksheet.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the protected workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
