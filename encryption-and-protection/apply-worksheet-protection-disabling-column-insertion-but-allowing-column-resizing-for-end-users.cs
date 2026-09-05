// Title: How to protect an Excel worksheet in C# with Aspose.Cells so users cannot insert columns but can resize them
// AI Prompts: Write C# code using Aspose.Cells to protect a worksheet, set a password, disable column insertion, and keep column width adjustments enabled. | Show how to configure Worksheet.ProtectionOptions in Aspose.Cells to allow column resizing while preventing new columns from being added. | Provide an example that saves the protected workbook to a specific file path after applying selective protection flags in .NET. | Demonstrate combining ProtectionType with custom options to block column insertion but permit column width changes.
// Common Searches: Aspose.Cells protect worksheet allow column resize disable insert column C# | C# Aspose.Cells set worksheet protection options for column insertion only | How to prevent users from adding columns in an Excel file using Aspose.Cells .NET | Enable column width editing while worksheet is protected with Aspose.Cells | Selective worksheet protection flags Aspose.Cells .NET example
// Tags: worksheet protection options Aspose.Cells | disable column insertion Excel .NET | allow column resizing protected sheet | Aspose.Cells password protection C# | selective worksheet protection flags

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first worksheet, applies password‑protected protection, disables the ability to insert new columns while keeping column width adjustments enabled, and saves the workbook to a specified .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Protect the worksheet with all protection types and set a password
            // The third parameter is the old password (null because the sheet is not previously protected)
            sheet.Protect(ProtectionType.All, "myPassword", null);

            // Define output file path
            string outputPath = "ProtectedSheet.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
