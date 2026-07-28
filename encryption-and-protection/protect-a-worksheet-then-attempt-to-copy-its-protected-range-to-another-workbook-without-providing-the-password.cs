// Title: Aspose.Cells .NET C# – Copy a password‑protected range to another workbook without supplying the password
// Description: This C# example creates a source workbook, adds data to cells A1:B2, defines an AllowEditRanges protected range with a password, and protects the worksheet with a different password. It then attempts to copy the protected range to a new workbook without providing the password, catches the expected exception, and saves both workbooks to show that the target remains empty.
// Keywords: Aspose.Cells copy protected range | C# protected range exception | worksheet protection Aspose.Cells | AllowEditRanges copy error | copy cells without password Aspose.Cells | Aspose.Cells .NET security | protected range copy failure
// Common Searches: copy password protected range Aspose.Cells .NET | exception when copying protected cells Aspose.Cells | how to bypass worksheet protection Aspose.Cells | Aspose.Cells copy range without password | C# Aspose.Cells protected range copy example
// Developer Intent: Attempt to copy a password‑protected worksheet range to another workbook without providing the password and observe the protection enforcement behavior.
// Use Cases: Demonstrate that copying a protected range without the correct password throws an exception. | Show how to check IsProtected and IsProtectedWithPassword flags before a copy operation. | Illustrate saving source and destination workbooks to verify that the destination remains unchanged after a failed copy.
// AI Prompts: Generate C# code using Aspose.Cells that copies a protected range to another workbook and handles the exception when the password is omitted. | Explain how to detect a password‑protected AllowEditRange and supply the password to enable a successful copy with Aspose.Cells. | Create a C# unit test that asserts an exception is thrown when copying a password‑protected range without providing its password using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsProtectedRangeCopyDemo
{
    // This C# example creates a source workbook, adds data to cells A1:B2, defines an AllowEditRanges protected range with a password, and protects the worksheet with a different password. It then attempts to copy the protected range to a new workbook without providing the password, catches the expected exception, and saves both workbooks to show that the target remains empty.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create source workbook ----------
                Workbook srcWb = new Workbook();
                Worksheet srcSheet = srcWb.Worksheets[0];

                // Add some data to the range that will be protected
                srcSheet.Cells["A1"].PutValue("Secret 1");
                srcSheet.Cells["B1"].PutValue("Secret 2");
                srcSheet.Cells["A2"].PutValue("Secret 3");
                srcSheet.Cells["B2"].PutValue("Secret 4");

                // Create a protected range (AllowEditRanges) covering A1:B2
                int rangeIndex = srcSheet.AllowEditRanges.Add("MyProtectedRange", 0, 0, 1, 1);
                ProtectedRange protectedRange = srcSheet.AllowEditRanges[rangeIndex];
                protectedRange.Password = "rangePwd"; // set password for the range

                // Protect the entire worksheet with a password
                srcSheet.Protect(ProtectionType.All, "sheetPwd", null);

                // Verify protection flags (optional)
                Console.WriteLine($"Worksheet protected: {srcSheet.IsProtected}");
                Console.WriteLine($"Range password protected: {protectedRange.IsProtectedWithPassword}");

                // Save source workbook (optional, just to visualize)
                srcWb.Save("SourceProtected.xlsx");

                // ---------- Create target workbook ----------
                Workbook tgtWb = new Workbook();
                Worksheet tgtSheet = tgtWb.Worksheets[0];

                // Attempt to copy the protected range without providing any password
                try
                {
                    // Define the source range object (same cells as the protected range)
                    Aspose.Cells.Range srcRange = srcSheet.Cells.CreateRange(0, 0, 2, 2); // rows 0-1, cols 0-1 (A1:B2)

                    // Define the destination range starting at A1 in target sheet
                    Aspose.Cells.Range tgtRange = tgtSheet.Cells.CreateRange("A1");

                    // Perform the copy operation
                    srcRange.Copy(tgtRange);
                    Console.WriteLine("Copy succeeded (unexpected).");
                }
                catch (Exception ex)
                {
                    // Expected: an exception because the source worksheet is protected with a password
                    Console.WriteLine("Copy failed as expected:");
                    Console.WriteLine(ex.Message);
                }

                // Save the target workbook to see the result (it will be empty)
                tgtWb.Save("TargetAfterCopyAttempt.xlsx");
            }
            catch (Exception e)
            {
                // General exception handling to prevent unexpected crashes
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(e.Message);
            }
        }
    }
}
