// Title: How to protect an Aspose.Cells worksheet with an empty password and verify its protection state in C#
// Common Searches: Aspose.Cells protect worksheet without specifying a password in C# | What does Worksheet.Protect do when the password is blank in Aspose.Cells? | Check if a worksheet is protected after calling Protect with no password using Aspose.Cells | Unprotect a worksheet that was protected with an empty password in Aspose.Cells for .NET | Save a workbook to a memory stream after protecting a worksheet with Aspose.Cells
// Tags: worksheet protect without password Aspose.Cells | Worksheet.IsProtected after Protect call | unprotect worksheet no password C# Aspose.Cells | protect all features worksheet Aspose.Cells | save workbook to memory stream Aspose.Cells | blank password protection behavior Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, protects the first worksheet using Worksheet.Protect(ProtectionType.All) with no password, prints the IsProtected flag and whether a password is stored, attempts to unprotect the sheet without a password, confirms the sheet is no longer protected, and finally saves the workbook to a memory stream.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook with a single worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "TestSheet";

            // Protect the worksheet without a password (use ProtectionType.All)
            sheet.Protect(ProtectionType.All);

            // Observe the protection state
            Console.WriteLine("After Protect without password:");
            Console.WriteLine($"IsProtected = {sheet.IsProtected}");
            Console.WriteLine($"Protection password is set? {(string.IsNullOrEmpty(sheet.Protection.Password) ? "No" : "Yes")}");

            // Attempt to unprotect (no password needed)
            try
            {
                sheet.Unprotect();
                Console.WriteLine("Unprotect succeeded without password.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unprotect threw an exception: {ex.Message}");
            }

            // Verify the worksheet is no longer protected
            Console.WriteLine($"IsProtected after Unprotect = {sheet.IsProtected}");

            // Save the workbook to a memory stream (no file I/O required for the test)
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to memory stream, length = {ms.Length} bytes.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
