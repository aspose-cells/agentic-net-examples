// Title: Attempt to copy a password‑protected worksheet range to another workbook without providing the password using Aspose.Cells for .NET
// Common Searches: Aspose.Cells copy range from a password protected worksheet without providing the password in C# | How to catch the exception when copying a protected range to another workbook using Aspose.Cells .NET | Copy cells from a worksheet protected with ProtectionType.All to a new workbook in C# | What error is thrown by Aspose.Cells when copying a protected range without a password | C# example of copying a protected sheet range to a different workbook using Aspose.Cells
// Tags: Aspose.Cells protected worksheet range copy | password protected sheet copy exception | range copy without password .NET | Aspose.Cells copy cells across workbooks | worksheet protection handling Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a source workbook, protects its first worksheet with a password, defines a range (A1:A2), and attempts to copy that protected range to a different workbook without supplying the password. The expected exception is caught, and both workbooks are saved.
class Program
{
    static void Main()
    {
        try
        {
            // -------------------- Create source workbook --------------------
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Name = "ProtectedSheet";

            // Fill some data in the sheet
            sourceSheet.Cells["A1"].PutValue("First");
            sourceSheet.Cells["A2"].PutValue("Second");

            // Protect the worksheet with a password (oldPassword is not required, pass empty string)
            sourceSheet.Protect(ProtectionType.All, "SecretPwd", string.Empty);

            // -------------------- Create destination workbook --------------------
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];
            destinationSheet.Name = "DestinationSheet";

            // -------------------- Attempt to copy protected range without password --------------------
            // Define the range to copy (A1:A2) from the protected sheet
            Aspose.Cells.Range protectedRange = sourceSheet.Cells.CreateRange("A1:A2");

            try
            {
                // Destination range starting at B1
                Aspose.Cells.Range destinationRange = destinationSheet.Cells.CreateRange("B1");

                // Try to copy the protected range to the destination sheet
                // No password is supplied, so this should fail or copy only values depending on protection settings
                protectedRange.Copy(destinationRange);
                Console.WriteLine("Range copied successfully (unexpected).");
            }
            catch (Exception ex)
            {
                // Expected path: copying fails because the source range is protected
                Console.WriteLine("Copy failed as expected: " + ex.Message);
            }

            // -------------------- Save workbooks (optional) --------------------
            sourceWorkbook.Save("SourceWorkbook.xlsx");
            destinationWorkbook.Save("DestinationWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            // General exception handling for unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
