// Title: How to lock rows 20‑25 and protect an Excel worksheet with a password using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to apply a locked style to rows 20 through 25 and then protect the worksheet with a password. | Show how to unlock all cells in a workbook, lock a specific row range, and save the file as an .xlsx using Aspose.Cells for .NET. | Write a C# snippet that creates a workbook, locks rows 20‑25 with a style, calls Worksheet.Protect with a password, and outputs ProtectedRows.xlsx.
// Common Searches: Aspose.Cells C# lock specific rows and set worksheet password | protect rows 20 to 25 in Excel file using Aspose.Cells .NET | apply locked style to entire row range with Aspose.Cells API | how to unlock all cells before protecting selected rows in Aspose.Cells | C# example for worksheet protection with password and row‑level locking
// Tags: lock rows Aspose.Cells | worksheet protect password Aspose.Cells | apply locked style to rows Aspose.Cells | unlock all cells Aspose.Cells | protect specific rows Excel Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace ProtectRowsExampleApp
{
    // The example creates a new workbook, unlocks every cell, applies a locked style to rows 20‑25, protects the worksheet with a password, and saves the result as ProtectedRows.xlsx.
    class ProtectRowsExample
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Unlock all cells in the worksheet first
                Style unlockedStyle = workbook.CreateStyle();
                unlockedStyle.IsLocked = false;
                StyleFlag unlockFlag = new StyleFlag { All = true };
                sheet.Cells.ApplyStyle(unlockedStyle, unlockFlag);

                // Define the password for protection
                string password = "MySecretPassword";

                // Create a style that locks cells
                Style lockedStyle = workbook.CreateStyle();
                lockedStyle.IsLocked = true;
                StyleFlag lockFlag = new StyleFlag { All = true };

                // Protect rows 20 through 25 (Excel rows are 1‑based)
                for (int rowIndex = 19; rowIndex <= 24; rowIndex++)
                {
                    // Apply the locked style to the entire row
                    sheet.Cells.Rows[rowIndex].ApplyStyle(lockedStyle, lockFlag);
                }

                // Protect the worksheet with the password
                sheet.Protect(ProtectionType.All, password, string.Empty);

                // Determine output path and ensure directory exists
                string outputPath = "ProtectedRows.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
