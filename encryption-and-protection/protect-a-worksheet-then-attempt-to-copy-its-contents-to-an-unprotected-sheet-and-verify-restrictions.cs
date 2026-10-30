// Title: How to password‑protect a worksheet, prevent edits, copy its contents to an unprotected sheet, and save the workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# protect worksheet with password and block cell edits | copy data from a protected worksheet to a new unprotected worksheet using Aspose.Cells | catch exception when trying to modify a protected cell in Aspose.Cells .NET | save workbook after copying a protected sheet to an unprotected sheet with Aspose.Cells | verify that worksheet protection remains enforced after sheet copy in C#
// Tags: Aspose.Cells worksheet password protection | duplicate protected sheet content to another worksheet | exception on modifying protected cell Aspose.Cells | saving workbook after sheet copy Aspose.Cells | worksheet protection verification .NET

using System;
using Aspose.Cells;

// Demonstrates applying password protection to a worksheet, handling the exception thrown when attempting to edit a protected cell, copying the entire protected sheet to a newly added unprotected worksheet, verifying the copied data, and finally saving the workbook.
class WorksheetProtectionDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and name it "ProtectedSheet"
            Worksheet protectedSheet = workbook.Worksheets[0];
            protectedSheet.Name = "ProtectedSheet";

            // Fill some data into the protected sheet
            protectedSheet.Cells["A1"].PutValue("Hello");
            protectedSheet.Cells["B1"].PutValue(123);
            protectedSheet.Cells["A2"].PutValue(DateTime.Now);
            protectedSheet.Cells["B2"].PutValue(3.14);

            // Protect the worksheet with a password
            protectedSheet.Protect(ProtectionType.All, "pwd", null);

            // Attempt to modify a cell in the protected sheet (should raise an exception)
            try
            {
                protectedSheet.Cells["A1"].PutValue("Modified");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Modification blocked as expected: " + ex.Message);
            }

            // Add a new unprotected worksheet
            Worksheet unprotectedSheet = workbook.Worksheets.Add("UnprotectedSheet");

            // Copy contents from the protected sheet to the unprotected sheet (copy all)
            protectedSheet.Copy(unprotectedSheet);

            // Verify that the copy succeeded by reading a value from the unprotected sheet
            string copiedValue = unprotectedSheet.Cells["A1"].StringValue;
            Console.WriteLine("Copied value in unprotected sheet: " + copiedValue);

            // Save the workbook to a file
            string outputPath = "WorksheetProtectionDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
