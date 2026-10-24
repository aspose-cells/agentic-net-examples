// Title: Create an XLS workbook, merge cells B1:B3, and protect the merged range with a password using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge a vertical range B1:B3 and set password protection | How to save a merged cell range as a password‑protected XLS file with Aspose.Cells | Apply sheet protection after merging cells B1 to B3 in a .NET workbook using Aspose.Cells
// Tags: merge cells B1:B3 Aspose.Cells | worksheet password protection Aspose.Cells | save workbook as XLS Aspose.Cells | protect merged range Aspose.Cells | cell merging with sheet protection .NET

using System;
using Aspose.Cells;

// The example creates a new workbook, merges cells B1:B3 on the first worksheet, applies full sheet protection with the password "MyPassword", and saves the file as an XLS workbook named MergedProtected.xls.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells B1:B3 (row 0, column 1, spanning 3 rows, 1 column)
            sheet.Cells.Merge(0, 1, 3, 1);

            // Protect the entire sheet with a password.
            // The third parameter (oldPassword) is not required here, so pass null.
            sheet.Protect(ProtectionType.All, "MyPassword", null);

            // Save the workbook; format is inferred from the file extension.
            workbook.Save("MergedProtected.xls");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
