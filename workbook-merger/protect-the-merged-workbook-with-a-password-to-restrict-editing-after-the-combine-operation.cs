// Title: Protect the structure of a merged workbook with a password and save as .xlsx using Aspose.Cells for .NET
// AI Prompts: Write C# code that merges worksheets, then applies structure protection with a password using Aspose.Cells and saves the result as an .xlsx file. | Show how to call Workbook.Protect with ProtectionType.Structure and a password after creating a merged workbook in C#. | Demonstrate saving a password‑protected merged workbook to disk with Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells protect merged workbook structure with password before saving | how to set ProtectionType.Structure with a password on a workbook after merging sheets using Aspose.Cells | saving a password‑protected Excel file after combining worksheets with Aspose.Cells for .NET | Aspose.Cells protect workbook after merge and export to .xlsx in C#
// Tags: Aspose.Cells workbook structure protection with password | merged workbook password protection C# | save protected workbook as xlsx Aspose.Cells | ProtectionType.Structure example Aspose.Cells | combine worksheets then protect Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // // Creates a workbook, adds sample data, protects its structure with a password via Workbook.Protect(ProtectionType.Structure, password), and saves the file as a password‑protected .xlsx using Aspose.Cells for .NET.
    public class ProtectMergedWorkbook
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add sample data.
                using (Workbook mergedWorkbook = new Workbook())
                {
                    Worksheet sheet = mergedWorkbook.Worksheets[0];
                    sheet.Cells["A1"].PutValue("Merged data");

                    // Protect the workbook structure with a password.
                    mergedWorkbook.Protect(ProtectionType.Structure, "MySecurePassword");

                    // Save the protected workbook.
                    mergedWorkbook.Save("MergedWorkbook_Protected.xlsx", SaveFormat.Xlsx);
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
            ProtectMergedWorkbook.Run();
        }
    }
}
