// Title: Aspose.Cells .NET – Copy a password‑protected worksheet without supplying the password
// Description: Demonstrates creating a workbook, protecting its first worksheet with a password, and then trying to copy that sheet to a new workbook without the password. The example catches the exception thrown by Aspose.Cells when protection blocks the copy operation.
// Keywords: Aspose.Cells copy protected worksheet | worksheet protection exception .NET | Copy without password Aspose.Cells | Worksheet.IsProtected check | Aspose.Cells ProtectionType.All
// Common Searches: copy password protected worksheet Aspose.Cells .NET | exception when copying protected sheet without password | how to detect protected worksheet before copy | Aspose.Cells copy fails due to protection | bypass worksheet protection during copy Aspose.Cells
// Developer Intent: Show that copying a password‑protected worksheet without the correct password triggers an exception in Aspose.Cells.
// Use Cases: Validate Worksheet.IsProtected before invoking Copy to avoid runtime errors. | Log the specific exception message when a protected sheet cannot be copied. | Unprotect the worksheet with the correct password or supply the password to enable a successful copy.
// AI Prompts: Provide Aspose.Cells .NET code that copies a protected worksheet after supplying the password. | Explain how to programmatically check Worksheet.IsProtected and conditionally copy or log a warning. | What exception type does Aspose.Cells throw when Copy is called on a password‑protected worksheet without a password?

using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetCopyDemo
{
    // Demonstrates creating a workbook, protecting its first worksheet with a password, and then trying to copy that sheet to a new workbook without the password. The example catches the exception thrown by Aspose.Cells when protection blocks the copy operation.
    class Program
    {
        static void Main()
        {
            // Create the source workbook and protect its first worksheet with a password
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Cells["A1"].PutValue("Sensitive Data");
            sourceSheet.Protect(ProtectionType.All, "SecretPwd", null);
            Console.WriteLine("Source worksheet protected: " + sourceSheet.IsProtected);

            // Save the source workbook (optional, just for inspection)
            sourceWorkbook.Save("SourceProtected.xlsx");

            // Create a new destination workbook
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            try
            {
                // Attempt to copy the protected worksheet without providing the password.
                // This will throw an exception because the source worksheet is password‑protected.
                sourceSheet.Copy(destinationSheet);
                Console.WriteLine("Copy succeeded unexpectedly.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Copy failed as expected: " + ex.Message);
            }

            // Save the destination workbook (will contain an empty sheet if copy failed)
            destinationWorkbook.Save("Destination.xlsx");
        }
    }
}
