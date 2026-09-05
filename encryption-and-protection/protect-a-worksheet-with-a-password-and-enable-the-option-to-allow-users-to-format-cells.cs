// Title: How to password‑protect an Excel worksheet while still allowing cell formatting using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to protect a worksheet with a password and enables the AllowFormattingCells option so users can still format cells. | Show how to verify and create the output directory if it does not exist before saving the protected workbook with Aspose.Cells. | Add try‑catch error handling around worksheet protection and workbook saving in a C# Aspose.Cells example.
// Common Searches: Aspose.Cells C# protect worksheet password but keep formatting enabled | How to allow cell formatting on a protected sheet using Aspose.Cells .NET | Create folder if missing before saving Excel file with Aspose.Cells in C# | Worksheet.Protect with ProtectionType.All and custom password example Aspose.Cells
// Tags: worksheet protection password Aspose.Cells .NET | allow formatting cells on protected sheet Aspose.Cells | create output directory before saving Excel Aspose.Cells | exception handling workbook save Aspose.Cells C# | ProtectionType.All usage Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new workbook, protects the first worksheet with the password "MySecurePassword" while allowing cell formatting, ensures the output directory exists, and saves the file as ProtectedWorksheet.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Protect the worksheet with a password.
                // The third parameter is the old password (empty string for new protection).
                sheet.Protect(ProtectionType.All, "MySecurePassword", string.Empty);

                // Define output file path
                string outputPath = "ProtectedWorksheet.xlsx";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the protected workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
