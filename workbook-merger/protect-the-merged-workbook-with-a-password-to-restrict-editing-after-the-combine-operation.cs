// Title: Merge two Excel workbooks and protect the combined file with a password using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells merge two .xlsx files and set workbook password protection | How to protect workbook structure after combining worksheets with Aspose.Cells .NET | Aspose.Cells protect merged workbook with password in C# example | AddCopy worksheets from another workbook and apply protection using Aspose.Cells | Protect merged Excel file from editing using Aspose.Cells Protect method
// Tags: Aspose.Cells workbook merging with password protection | Workbook.Protect method with ProtectionType.All in C# | AddCopy worksheets from another workbook Aspose.Cells | combine multiple .xlsx files using Aspose.Cells | protect merged Excel file structure .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads two Excel workbooks, copies each worksheet from the second workbook into the first using AddCopy, applies full workbook protection with a password via Protect(ProtectionType.All, password), and saves the result as a protected merged file while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string firstFile = "Book1.xlsx";
            const string secondFile = "Book2.xlsx";
            const string outputFile = "MergedProtected.xlsx";
            const string password = "YourPasswordHere";

            // Verify input files exist
            if (!File.Exists(firstFile))
                throw new FileNotFoundException($"Input file not found: {firstFile}");
            if (!File.Exists(secondFile))
                throw new FileNotFoundException($"Input file not found: {secondFile}");

            // Load the first workbook (the target workbook)
            Workbook mergedWorkbook = new Workbook(firstFile);

            // Load the second workbook to be merged
            Workbook workbookToMerge = new Workbook(secondFile);

            // Copy each worksheet from the second workbook into the first workbook
            foreach (Worksheet ws in workbookToMerge.Worksheets)
            {
                // AddCopy expects the source worksheet name, not the Worksheet object itself
                mergedWorkbook.Worksheets.AddCopy(ws.Name);
            }

            // Protect the merged workbook (structure, windows, etc.)
            mergedWorkbook.Protect(ProtectionType.All, password);

            // Save the protected workbook
            mergedWorkbook.Save(outputFile);
            Console.WriteLine($"Workbooks merged and saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
