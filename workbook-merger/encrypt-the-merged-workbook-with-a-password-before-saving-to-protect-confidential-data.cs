// Title: Encrypt a merged Excel workbook with a password using Aspose.Cells for .NET (C#)
// Common Searches: how to password protect a workbook after merging sheets with Aspose.Cells in C# | Aspose.Cells C# encrypt merged workbook before saving | set workbook password for merged Excel file using Aspose.Cells .NET | protect confidential data in merged workbook Aspose.Cells C# example | apply password protection to a workbook created by merging multiple Excel files with Aspose.Cells
// Tags: Aspose.Cells workbook password encryption | C# merge worksheets and secure workbook | Workbook.Settings.Password usage Aspose.Cells | SaveFormat.Xlsx encrypted output Aspose.Cells | confidential Excel merge protection .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads two source Excel files, copies their first worksheets into a new workbook, removes the default empty sheet, sets a password via the Workbook.Settings.Password property, and saves the merged workbook as an encrypted MergedEncrypted.xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook that will contain the merged data
            Workbook mergedWorkbook = new Workbook();

            // Paths to source workbooks
            string srcPath1 = "Source1.xlsx";
            string srcPath2 = "Source2.xlsx";

            // Load source workbooks if they exist
            Workbook src1 = null;
            Workbook src2 = null;

            if (File.Exists(srcPath1))
            {
                src1 = new Workbook(srcPath1);
            }
            else
            {
                Console.WriteLine($"File not found: {srcPath1}");
            }

            if (File.Exists(srcPath2))
            {
                src2 = new Workbook(srcPath2);
            }
            else
            {
                Console.WriteLine($"File not found: {srcPath2}");
            }

            // Copy worksheets from source workbooks into the merged workbook
            if (src1 != null && src1.Worksheets.Count > 0)
            {
                mergedWorkbook.Worksheets.AddCopy(src1.Worksheets[0].Name);
            }

            if (src2 != null && src2.Worksheets.Count > 0)
            {
                mergedWorkbook.Worksheets.AddCopy(src2.Worksheets[0].Name);
            }

            // Remove the initially created empty sheet (optional)
            if (mergedWorkbook.Worksheets.Count > 0)
            {
                mergedWorkbook.Worksheets.RemoveAt(0);
            }

            // Encrypt the workbook with a password
            mergedWorkbook.Settings.Password = "Confidential123";

            // Save the encrypted workbook
            string outputPath = "MergedEncrypted.xlsx";
            mergedWorkbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Merged workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
