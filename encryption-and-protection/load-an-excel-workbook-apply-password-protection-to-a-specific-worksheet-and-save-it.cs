// Title: How to password‑protect a single worksheet in an Excel file using Aspose.Cells for .NET (C#) and save it
// AI Prompts: Load an existing Excel file (or create a new workbook) with Aspose.Cells, apply a password to the 'Sheet1' worksheet using ProtectionType.All, and write the protected workbook to a new file. | Write C# code that checks for 'input.xlsx', opens it with Aspose.Cells, secures a specific sheet with a custom password, and saves the result as 'output.xlsx'.
// Common Searches: Aspose.Cells example to protect only Sheet1 with a password in C# | C# protect a single worksheet using ProtectionType.All in Aspose.Cells | How to save an Excel workbook with a password‑protected sheet using Aspose.Cells for .NET | Conditional workbook loading and sheet protection with Aspose.Cells | Set worksheet password without encrypting the whole file Aspose.Cells
// Tags: worksheet.Protect password Aspose.Cells C# | ProtectionType.All on Excel sheet | conditional workbook load Aspose.Cells | save protected workbook to .xlsx | single sheet password protection .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an existing 'input.xlsx' workbook (or creates a new one if the file is missing), selects the worksheet named 'Sheet1' (or the first sheet), applies password protection with worksheet.Protect using ProtectionType.All and a custom password, and then saves the protected workbook as 'output.xlsx', handling any errors that may occur.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    workbook.Worksheets[0].Name = "Sheet1";
                }

                // Get the worksheet by name; fallback to the first sheet if not found
                Worksheet worksheet = workbook.Worksheets["Sheet1"] ?? workbook.Worksheets[0];

                // Protect the worksheet (oldPassword is empty because the sheet is not previously protected)
                worksheet.Protect(ProtectionType.All, "MySecurePassword", string.Empty);

                // Save the protected workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
