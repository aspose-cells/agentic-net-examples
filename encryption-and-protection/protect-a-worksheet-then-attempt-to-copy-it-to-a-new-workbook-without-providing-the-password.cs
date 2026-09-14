// Title: Copy a password-protected worksheet to a new workbook using Aspose.Cells for .NET and capture the missing-password exception
// AI Prompts: Generate C# code that protects a worksheet with a password, attempts to copy it to another workbook using Worksheets.AddCopy, and catches the exception thrown when the password is not supplied. | Show how to detect and log the error message returned by Aspose.Cells when copying a protected sheet without providing its password. | Provide an example that saves both the original protected workbook and the destination workbook after a failed copy operation in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells .NET copy protected sheet without password exception | How to handle Worksheets.AddCopy error for password-protected worksheet in C# | Saving source and destination workbooks after failed copy of protected sheet Aspose.Cells | C# Aspose.Cells copy worksheet protected by password and get error message | What exception is thrown when copying a password-protected worksheet using Aspose.Cells
// Tags: Aspose.Cells worksheet copy protected exception | Worksheets.AddCopy password protection .NET | C# handling missing password during sheet copy | save workbooks after copy failure Aspose.Cells | protect worksheet with password Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, protects its first worksheet with a password, then tries to copy that sheet to a new workbook using Worksheets.AddCopy without supplying the password, catches the resulting exception, and finally saves both workbooks.
class Program
{
    static void Main()
    {
        // Create source workbook with one worksheet
        Workbook srcWb = new Workbook();
        Worksheet srcWs = srcWb.Worksheets[0];
        srcWs.Name = "ProtectedSheet";

        // Add some sample data
        srcWs.Cells["A1"].PutValue("Sample Text");
        srcWs.Cells["A2"].PutValue(42);

        // Protect the worksheet with a password (oldPassword is not required for new protection)
        srcWs.Protect(ProtectionType.All, "SecretPwd", string.Empty);

        // Create destination workbook (initially empty)
        Workbook destWb = new Workbook();

        try
        {
            // Attempt to copy the protected worksheet without providing the password.
            // This will raise an exception because the source worksheet is protected.
            destWb.Worksheets.AddCopy(srcWs.Name);
        }
        catch (Exception ex)
        {
            // Output the error message to indicate the copy failed
            Console.WriteLine("Copy operation failed: " + ex.Message);
        }

        try
        {
            // Save both workbooks for verification
            srcWb.Save("ProtectedSource.xlsx");
            destWb.Save("Destination.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error saving workbooks: " + ex.Message);
        }
    }
}
