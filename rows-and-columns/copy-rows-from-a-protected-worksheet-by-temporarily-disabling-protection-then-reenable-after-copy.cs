using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // -------------------------------------------------
        // Prepare the source worksheet (protected)
        // -------------------------------------------------
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "Source";

        // Populate some sample rows in the source sheet
        sourceSheet.Cells["A1"].PutValue("Source Row 1");
        sourceSheet.Cells["A2"].PutValue("Source Row 2");
        sourceSheet.Cells["A3"].PutValue("Source Row 3");

        // Protect the source worksheet with a password
        string password = "pwd123";
        sourceSheet.Protect(ProtectionType.All, password, null);

        // -------------------------------------------------
        // Prepare the destination worksheet (unprotected)
        // -------------------------------------------------
        Worksheet destSheet = workbook.Worksheets.Add("Destination");

        // -------------------------------------------------
        // Temporarily disable protection, copy rows, then re‑enable protection
        // -------------------------------------------------
        // Unprotect the source sheet using the password
        sourceSheet.Unprotect(password);

        // Copy all rows from the source sheet to the destination sheet
        // sourceRowIndex = 0 (first row), destinationRowIndex = 0, rowNumber = total rows in source
        int rowsToCopy = sourceSheet.Cells.MaxDisplayRange.RowCount;
        destSheet.Cells.CopyRows(sourceSheet.Cells, 0, 0, rowsToCopy);

        // Re‑protect the source sheet with the same password
        sourceSheet.Protect(ProtectionType.All, password, null);

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("CopyRowsFromProtectedWorksheet.xlsx");
    }
}