// Title: Lock rows 20‑25 and apply worksheet password protection with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# lock rows 20 to 25 and set worksheet password | protect specific rows in an Excel file using Aspose.Cells .NET | how to use EntireRow property to lock rows with Aspose.Cells | C# example for setting IsLocked style on multiple rows in Aspose.Cells | apply password protection to a worksheet after locking rows in Aspose.Cells
// Tags: lock rows Aspose.Cells C# | worksheet password protection Aspose.Cells | EntireRow property usage Aspose.Cells | row IsLocked style Aspose.Cells | protect specific rows Excel Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, fills column A with sample data, locks rows 20‑25 by setting the IsLocked style on each row, protects the worksheet with the password "MySecretPassword", and saves the file as ProtectedRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data for demonstration
            for (int r = 0; r < 30; r++)
            {
                sheet.Cells[r, 0].PutValue($"Row {r + 1}");
            }

            // Lock rows 20 through 25 (zero‑based indices 19‑24)
            for (int rowIndex = 19; rowIndex <= 24; rowIndex++)
            {
                Row row = sheet.Cells.Rows[rowIndex];
                Style rowStyle = row.GetStyle();
                rowStyle.IsLocked = true;          // Ensure the row is locked
                row.SetStyle(rowStyle);
            }

            // Protect the worksheet with a password (protect all options)
            // The third parameter is the old password; pass null when not required
            sheet.Protect(ProtectionType.All, "MySecretPassword", null);

            // Determine output path and ensure directory exists
            string outputPath = "ProtectedRows.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
