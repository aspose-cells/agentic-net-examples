// Title: Copy a password‑protected worksheet to a new workbook without providing the password using Aspose.Cells for .NET
// Common Searches: Aspose.Cells copy worksheet that is protected with a password without needing the password in C# | How to duplicate a protected sheet to another workbook using Aspose.Cells .NET API | Can AddCopy copy a worksheet when ProtectionType.All is applied in Aspose.Cells | Copy protected worksheet to new workbook without unprotecting in C# Aspose.Cells | Worksheet protection impact on copying sheets with Aspose.Cells for .NET
// Tags: AddCopy protected sheet Aspose.Cells | worksheet protection copy behavior Aspose.Cells | transfer worksheet between workbooks .NET | ProtectionType.All copy support Aspose.Cells | copy sheet without unprotecting Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a source workbook, protects its first worksheet with a password using ProtectionType.All, then creates a destination workbook and copies the protected worksheet by name with Worksheets.AddCopy, which does not require the password, and finally saves both workbooks.
class Program
{
    static void Main()
    {
        try
        {
            // Create a source workbook and add some data
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Name = "ProtectedSheet";
            sourceSheet.Cells["A1"].PutValue("Sample Data");

            // Protect the worksheet with a password (protect all options)
            // The third parameter is the old password; null is used when setting a new password.
            sourceSheet.Protect(ProtectionType.All, "myPassword", null);

            // Create a destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Copy the protected worksheet by name (password is not required for copying)
            destinationWorkbook.Worksheets.AddCopy(sourceSheet.Name);
            Console.WriteLine("Worksheet copied successfully.");

            // Save the workbooks (optional) – ensure the directory is writable
            sourceWorkbook.Save("SourceWorkbook.xlsx");
            destinationWorkbook.Save("DestinationWorkbook.xlsx");
            Console.WriteLine("Workbooks saved successfully.");
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during processing
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
