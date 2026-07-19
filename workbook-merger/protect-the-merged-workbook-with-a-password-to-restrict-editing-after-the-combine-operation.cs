// Title: Password‑protect a merged Excel workbook’s structure with Aspose.Cells for .NET (C#)
// Description: The example opens a merged .xlsx (or creates a placeholder), applies structure protection via Workbook.Protect with a user‑defined password, and writes the secured file to a new location.
// Keywords: Aspose.Cells | C# | Workbook.Protect | ProtectionType.Structure | password protect Excel | merged workbook | protect workbook structure | Aspose.Cells .NET example | Excel file security | save protected workbook
// Common Searches: C# protect merged Excel workbook with password Aspose.Cells | How to lock workbook structure after combining files using Aspose.Cells .NET | Aspose.Cells protect workbook example | Password protect Excel file programmatically C# | Aspose.Cells protect merged workbook structure
// Developer Intent: Add a password that locks the structure of a merged workbook so sheets cannot be added, removed, renamed, or moved after the combine operation.
// Use Cases: Secure a consolidated financial report before sending it to auditors. | Prevent accidental edits to a merged template shared across multiple departments. | Generate a read‑only version of a combined workbook for external partners while retaining an editable copy for internal use.
// AI Prompts: Write C# code using Aspose.Cells to open an existing merged .xlsx, protect its structure with a password, and save the result as a new file. | Show how to catch and log exceptions when calling Workbook.Protect in Aspose.Cells for .NET. | Explain the difference between Workbook.Protect(ProtectionType.Structure) and Worksheet.Protect for cell‑level security, and provide code snippets for each.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example opens a merged .xlsx (or creates a placeholder), applies structure protection via Workbook.Protect with a user‑defined password, and writes the secured file to a new location.
    class ProtectMergedWorkbook
    {
        public static void Run()
        {
            const string inputPath = "merged.xlsx";
            const string outputPath = "merged_protected.xlsx";
            const string password = "MySecurePassword";

            try
            {
                // Load existing workbook or create a placeholder if missing
                Workbook mergedWorkbook;
                if (File.Exists(inputPath))
                {
                    mergedWorkbook = new Workbook(inputPath);
                }
                else
                {
                    mergedWorkbook = new Workbook();
                    mergedWorkbook.Save(inputPath, SaveFormat.Xlsx);
                }

                // Protect the workbook's structure
                mergedWorkbook.Protect(ProtectionType.Structure, password);

                // Save the protected workbook
                mergedWorkbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Protected workbook saved to '{outputPath}'.");

                mergedWorkbook.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ProtectMergedWorkbook.Run();
        }
    }
}
