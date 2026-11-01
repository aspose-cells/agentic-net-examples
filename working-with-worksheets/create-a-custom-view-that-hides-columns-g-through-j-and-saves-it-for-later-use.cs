// Title: Create a custom view that hides columns G‑J in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells hide columns G through J and keep view when reopening in C# | How to programmatically hide a column range and save as a custom view with Aspose.Cells | C# code to hide columns G-J in an Excel file and preserve hidden state | Saving a workbook with hidden columns using the Aspose.Cells .NET API | Create a custom view with hidden columns in Aspose.Cells for .NET
// Tags: hide column range Aspose.Cells C# | custom view creation Aspose.Cells .NET | save workbook with hidden columns Aspose.Cells | Excel column visibility manipulation Aspose.Cells | programmatic column hiding Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, hides columns G through J (indices 6‑9) on the first worksheet using Cells.HideColumn, and saves the file as CustomView.xlsx, effectively storing a custom view with those columns hidden.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Hide columns G (index 6) through J (index 9)
            for (int col = 6; col <= 9; col++)
            {
                worksheet.Cells.HideColumn(col);
            }

            // Save the workbook
            string outputPath = "CustomView.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
