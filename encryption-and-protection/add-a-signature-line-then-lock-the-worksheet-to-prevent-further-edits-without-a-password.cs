// Title: Create a signature line placeholder shape and password‑protect an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to insert a signature line shape in an Excel file with Aspose.Cells C# | Aspose.Cells C# protect worksheet with password example | add rectangle shape as placeholder and lock sheet using Aspose.Cells | save a password‑protected workbook with Aspose.Cells .NET | C# code to create signature line and protect Excel sheet
// Tags: add signature line shape Aspose.Cells | worksheet password protection Aspose.Cells | create rectangle shape in Excel C# | save protected Excel workbook Aspose.Cells | protect worksheet with ProtectionType.All C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, adds a rectangular shape labeled as a signature line placeholder on the first worksheet, applies full protection with a specified password, ensures the output directory exists, and saves the file as SignedAndProtected.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a placeholder shape (rectangle) to simulate a signature line
            // Parameters: type, upper left row, upper left column, lower right row, lower right column, height (pixels), width (pixels)
            Shape placeholder = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                1, // upper left row (zero‑based)
                1, // upper left column
                2, // lower right row
                2, // lower right column
                30, // height in pixels
                200 // width in pixels
            );
            placeholder.Text = "Signature Line Placeholder";

            // Protect the worksheet with a password
            worksheet.Protect(ProtectionType.All, "MySecurePassword", string.Empty);

            // Define output path and ensure the directory exists
            string outputPath = "SignedAndProtected.xlsx";
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
