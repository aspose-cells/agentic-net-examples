// Title: Hide columns G through J in an Aspose.Cells worksheet and save the workbook (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to hide columns G‑J (indices 6‑9) in the first worksheet and then saves the workbook as an .xlsx file. | Show how to programmatically hide a specific range of columns in an Aspose.Cells workbook and persist the changes to disk.
// Common Searches: Aspose.Cells C# hide columns G to J and save workbook | How to hide a range of columns in Aspose.Cells .NET | C# Aspose.Cells hide columns by index example | Saving hidden column settings with Aspose.Cells | Aspose.Cells hide columns programmatically and export to Excel
// Tags: hide columns Aspose.Cells C# | Aspose.Cells column visibility range | Aspose.Cells save workbook with hidden columns | Aspose.Cells hide column indices | Aspose.Cells worksheet column hide example

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a new Workbook, accesses the first worksheet, hides columns G through J using Cells.HideColumns(6, 4), and saves the file as CustomViewHideColumns.xlsx.
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

            // Hide columns G (index 6) through J (index 9)
            // HideColumns(startColumnIndex, totalColumnsToHide)
            sheet.Cells.HideColumns(6, 4); // hides columns 6,7,8,9

            // Save the workbook
            string outputPath = "CustomViewHideColumns.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
