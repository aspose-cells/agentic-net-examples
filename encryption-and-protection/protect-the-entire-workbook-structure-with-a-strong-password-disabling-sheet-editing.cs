// Title: C# – Protect an Aspose.Cells Workbook Structure and All Worksheets with a Strong Password
// Description: Demonstrates how to create a new Workbook with Aspose.Cells for .NET, apply structure protection to block adding, deleting, or renaming sheets, protect every worksheet with full protection using the same strong password, ensure the output folder exists, and save the file as an XLSX document.
// Keywords: Aspose.Cells protect workbook structure C# | Excel workbook password protection .NET | C# protect all worksheets Aspose.Cells | structure protection Aspose.Cells | strong password Excel file Aspose | disable sheet editing Aspose.Cells | save protected workbook C# | Aspose.Cells workbook.Protect example | worksheet.Protect all types C#
// Common Searches: How to lock workbook structure with Aspose.Cells C# | Apply a strong password to an entire Excel file using Aspose.Cells | Prevent adding or deleting sheets in Aspose.Cells .NET | Protect all worksheets with one password in Aspose.Cells | C# code to protect Excel workbook and worksheets
// Developer Intent: The developer needs to secure the workbook’s layout and every sheet with a single strong password so that users cannot modify structure or cell content.
// Use Cases: Distribute a read‑only template where users can only enter data, not change sheet order or add new sheets. | Share confidential financial reports that must retain their original layout and be tamper‑proof. | Generate compliance‑required workbooks that forbid deleting, renaming, or editing any worksheet.
// AI Prompts: Show me C# code that protects an Aspose.Cells workbook’s structure and all its worksheets with one strong password. | Provide an example that creates the output directory if it does not exist before saving the protected workbook. | Explain how to programmatically unprotect the workbook and its worksheets using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a new Workbook with Aspose.Cells for .NET, apply structure protection to block adding, deleting, or renaming sheets, protect every worksheet with full protection using the same strong password, ensure the output folder exists, and save the file as an XLSX document.
public class ProtectWorkbookDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook instance
            using (Workbook workbook = new Workbook())
            {
                // Define a strong password for protection
                string strongPassword = "Str0ngP@ssw0rd!2026";

                // Protect the workbook structure (prevents adding, deleting, renaming sheets)
                workbook.Protect(ProtectionType.Structure, strongPassword);

                // Additionally protect each worksheet to disable cell editing
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Protect the worksheet with all protection types using the same password
                    sheet.Protect(ProtectionType.All, strongPassword, null);
                }

                // Ensure the output directory exists
                string outputPath = "ProtectedWorkbook.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the protected workbook to a file
                workbook.Save(outputPath, SaveFormat.Xlsx);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        ProtectWorkbookDemo.Run();
    }
}
