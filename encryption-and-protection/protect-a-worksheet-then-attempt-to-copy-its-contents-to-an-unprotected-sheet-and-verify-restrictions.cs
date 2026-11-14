// Title: Protect a worksheet, copy its rows to an unprotected sheet, and verify protection behavior with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# protect worksheet with ProtectionType.All and copy rows to another sheet | How to copy rows from a protected worksheet to an unprotected worksheet using Aspose.Cells | Check IsProtected flag after copying data between worksheets in Aspose.Cells | What exception is thrown when modifying a protected worksheet in Aspose.Cells C# | Save workbook after applying worksheet protection with Aspose.Cells for .NET
// Tags: ProtectionType.All worksheet lock Aspose.Cells | CopyRows method protected sheet to target sheet C# | IsProtected status check Aspose.Cells | Modification attempt exception on locked worksheet Aspose.Cells | Workbook save post‑protection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, locks the first worksheet using ProtectionType.All, copies all rows to a newly added unprotected worksheet via Cells.CopyRows, prints the IsProtected status of both sheets, attempts to modify a cell in the locked sheet to capture the expected exception, and finally saves the workbook.
class WorksheetProtectionDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some sample data
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "ProtectedSheet";
            sourceSheet.Cells["A1"].PutValue("Header");
            sourceSheet.Cells["A2"].PutValue(123);
            sourceSheet.Cells["B2"].PutValue(456);

            // Protect the source worksheet (all protection options)
            sourceSheet.Protect(ProtectionType.All);

            // Verify protection status
            Console.WriteLine($"Is '{sourceSheet.Name}' protected? {sourceSheet.IsProtected}");

            // Add a new unprotected worksheet to copy data into
            Worksheet targetSheet = workbook.Worksheets.Add("UnprotectedCopy");

            // Copy rows from the protected sheet to the unprotected sheet
            // Note: Use the overload that takes source Cells as the first argument
            try
            {
                int totalRows = sourceSheet.Cells.MaxDisplayRange.RowCount;
                targetSheet.Cells.CopyRows(sourceSheet.Cells, 0, totalRows, 0);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during row copy: {ex.Message}");
            }

            // Verify that the target sheet is not protected
            Console.WriteLine($"Is '{targetSheet.Name}' protected? {targetSheet.IsProtected}");

            // Attempt to modify a cell in the protected sheet (should raise an exception)
            try
            {
                sourceSheet.Cells["A2"].PutValue(999);
                Console.WriteLine("Modification succeeded unexpectedly.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Expected exception when modifying protected sheet: {ex.Message}");
            }

            // Save the workbook
            string outputPath = "WorksheetProtectionDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
