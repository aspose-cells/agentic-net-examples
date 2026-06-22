using System;
using System.IO;
using Aspose.Cells;

class WorksheetProtectionDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some data
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Cells["A1"].PutValue("Protected Data");
            sourceSheet.Cells["B2"].PutValue(123);

            // Protect the source worksheet with a password
            sourceSheet.Protect(ProtectionType.All, "pwd123", null);
            Console.WriteLine("Source worksheet protected: " + sourceSheet.IsProtected);

            // Add a new unprotected worksheet to receive the copy
            int targetIndex = workbook.Worksheets.Add();
            Worksheet targetSheet = workbook.Worksheets[targetIndex];
            targetSheet.Name = "CopyTarget";

            // Copy rows from the protected sheet to the unprotected sheet
            // Correct overload: targetSheet.Cells.CopyRows(sourceSheet.Cells, sourceRowIndex, totalRows, targetRowIndex)
            targetSheet.Cells.CopyRows(sourceSheet.Cells, 0, 2, 0);
            Console.WriteLine("Rows copied from protected sheet to unprotected sheet.");

            // Verify that the target sheet remains unprotected
            Console.WriteLine("Target worksheet protected: " + targetSheet.IsProtected);

            // Attempt to modify a cell in the protected sheet without unprotecting
            try
            {
                sourceSheet.Cells["A1"].PutValue("Attempted Change");
                Console.WriteLine("Modified protected sheet without unprotecting (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error modifying protected sheet: " + ex.Message);
            }

            // Unprotect the source sheet using the correct password
            sourceSheet.Unprotect("pwd123");
            Console.WriteLine("Source worksheet after unprotect: " + sourceSheet.IsProtected);

            // Now modification should succeed
            sourceSheet.Cells["A1"].PutValue("Modified after unprotect");
            Console.WriteLine("Cell modified after unprotect.");

            // Save the workbook (ensure the directory exists)
            string outputPath = "WorksheetProtectionDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + Path.GetFullPath(outputPath));
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error saving workbook: " + ex.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}