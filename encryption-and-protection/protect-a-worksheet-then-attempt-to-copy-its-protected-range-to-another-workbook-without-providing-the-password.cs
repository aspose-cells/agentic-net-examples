// Title: Copy a protected worksheet range to another workbook without a password using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells copy range from a protected worksheet without password C# | How to copy cells from a locked sheet to a new workbook using Aspose.Cells | Exception thrown when copying protected range with Aspose.Cells .NET | Copy protected worksheet data to another workbook example Aspose.Cells
// Tags: copy protected range Aspose.Cells .NET | worksheet protection password handling Aspose.Cells | range copy exception Aspose.Cells | protected worksheet to another workbook Aspose.Cells | Aspose.Cells copy range without password

using System;
using Aspose.Cells;

// The example creates a source workbook, protects its first worksheet with a password, and then attempts to copy the A1:B2 range to a destination workbook without supplying the password. The operation fails, the expected exception is caught, and both workbooks are saved.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create source workbook ----------
            Workbook srcWb = new Workbook();
            Worksheet srcWs = srcWb.Worksheets[0];
            srcWs.Name = "ProtectedSheet";

            // Populate some data in the source worksheet
            srcWs.Cells["A1"].PutValue("Header");
            srcWs.Cells["A2"].PutValue(123);
            srcWs.Cells["B2"].PutValue(456);

            // Protect the worksheet with a password
            string password = "Secret123";
            srcWs.Protect(ProtectionType.All);          // enable protection
            srcWs.Protection.Password = password;       // set password

            // Save the source workbook (optional, just for verification)
            srcWb.Save("Source.xlsx");

            // ---------- Create destination workbook ----------
            Workbook destWb = new Workbook();
            Worksheet destWs = destWb.Worksheets[0];
            destWs.Name = "DestinationSheet";

            // ---------- Attempt to copy protected range without password ----------
            try
            {
                // Define the range to copy from the protected worksheet
                Aspose.Cells.Range srcRange = srcWs.Cells.CreateRange("A1:B2");

                // Define the target range in the destination worksheet
                Aspose.Cells.Range destRange = destWs.Cells.CreateRange("C3:D4");

                // This operation will fail because the source worksheet is protected
                // and no password is supplied.
                srcRange.Copy(destRange);

                Console.WriteLine("Copy succeeded unexpectedly.");
            }
            catch (Exception ex)
            {
                // Expected exception: cannot copy from a protected worksheet without password
                Console.WriteLine("Copy failed as expected: " + ex.Message);
            }

            // Save the destination workbook
            destWb.Save("Destination.xlsx");
        }
        catch (Exception e)
        {
            Console.WriteLine("An unexpected error occurred: " + e.Message);
        }
    }
}
