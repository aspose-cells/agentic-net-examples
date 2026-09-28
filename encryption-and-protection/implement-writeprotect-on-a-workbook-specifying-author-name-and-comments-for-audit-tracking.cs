// Title: Apply write‑protection with password and author metadata to an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an existing .xlsx file (or creates a new workbook if missing), sets Workbook.Settings.WriteProtection.Password and .Author, then saves the workbook with write‑protection applied. | Show how to use Aspose.Cells WorkbookSettings.WriteProtection to embed author information for audit tracking before saving a protected Excel file.
// Common Searches: Aspose.Cells C# how to add write protection password and author to an Excel file | set workbook author for write‑protected Excel using Aspose.Cells .NET | protect existing or new workbook with password and audit info in Aspose.Cells | C# Aspose.Cells write protection with custom author metadata example | save Excel workbook with write protection and author comment using Aspose.Cells
// Tags: Workbook.WriteProtection password Aspose.Cells | Aspose.Cells write protection author metadata | C# set write protection on .xlsx file | initialize workbook then apply write protection | audit trail write protection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWriteProtectionDemo
{
    // The sample loads an existing Excel file if present, otherwise creates a new workbook, configures Workbook.Settings.WriteProtection with a password and author for audit tracking, ensures the output directory exists, and saves the protected workbook as ProtectedWorkbook.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to an optional template workbook.
                string templatePath = "InputFile.xlsx";

                Workbook workbook;

                // Load existing workbook if the file exists; otherwise create a new one.
                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    workbook = new Workbook(); // creates a new empty workbook
                }

                // Configure write‑protection details (password, author).
                // The WriteProtectionInfo object is obtained from WorkbookSettings.
                workbook.Settings.WriteProtection.Password = "StrongPassword123";
                workbook.Settings.WriteProtection.Author = "John Doe";

                // Save the protected workbook.
                string outputPath = "ProtectedWorkbook.xlsx";

                // Ensure the directory for the output file exists.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log or display the error details.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
