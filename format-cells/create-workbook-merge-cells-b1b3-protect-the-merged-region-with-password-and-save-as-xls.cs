// Title: Merge B1:B3, protect with password, and save as XLS using Aspose.Cells for .NET
// Description: C# example that creates a new workbook, merges the range B1:B3 on the first worksheet, writes a value into the merged cell, applies full worksheet protection with a password, and saves the file as MergedProtected.xls.
// Keywords: Aspose.Cells | C# | merge cells B1:B3 | worksheet protection | password protected XLS | cells.Merge | worksheet.Protect | save workbook as XLS | merged region security
// Common Searches: Aspose.Cells merge cells and protect with password | C# protect merged cells in Excel file | Save protected merged range as .xls using Aspose.Cells | How to lock a merged cell in Aspose.Cells .NET | Worksheet.Protect password example Aspose.Cells
// Developer Intent: Create a .xls workbook, merge B1:B3, apply password protection, and save the file.
// Use Cases: Lock a title header that spans B1:B3 before distributing a template to users. | Secure a merged header row in a financial report while keeping the rest of the sheet editable. | Generate an XLS file for external partners where the merged cell is password‑protected to prevent accidental changes.
// AI Prompts: Generate C# code with Aspose.Cells to merge cells B1:B3, set a value, protect the worksheet using password "myPassword", and save as "MergedProtected.xls". | Explain how worksheet protection in Aspose.Cells impacts merged cells and how to enable it programmatically. | Show how to protect only a specific merged region while leaving other cells editable in an Aspose.Cells .NET workbook.

using System;
using System.IO;
using Aspose.Cells;

// C# example that creates a new workbook, merges the range B1:B3 on the first worksheet, writes a value into the merged cell, applies full worksheet protection with a password, and saves the file as MergedProtected.xls.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge cells B1:B3 (rows 0-2, column 1)
            cells.Merge(0, 1, 3, 1);

            // Set a value in the merged cell
            cells[0, 1].PutValue("Protected Merged Cell");

            // Protect the worksheet (including the merged region) with a password
            worksheet.Protect(ProtectionType.All, "myPassword", null);

            // Define output file path
            string outputPath = "MergedProtected.xls";

            // Save the workbook; format is inferred from the file extension
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
