// Title: Aspose.Cells for .NET – Protect a Worksheet, Copy Its Data to an Unprotected Sheet, and Validate Protection (C#)
// Description: This C# example demonstrates how to protect a worksheet with a password using Aspose.Cells, confirm that write attempts raise an exception, copy the protected sheet's content to a new unprotected worksheet, check the IsProtected flag on both sheets, edit the copied sheet successfully, and save the workbook.
// Keywords: Aspose.Cells protect worksheet C# | copy protected sheet Aspose.Cells | Worksheet.IsProtected property | worksheet protection exception .NET | Aspose.Cells copy without protection | C# Excel protection example | Aspose.Cells worksheet copy behavior
// Common Searches: how to protect a worksheet with Aspose.Cells .NET | copy protected worksheet to editable sheet Aspose.Cells | exception when writing to a protected sheet Aspose.Cells | does Worksheet.Copy keep protection settings | verify IsProtected after copying worksheet Aspose.Cells
// Developer Intent: The developer needs to lock a worksheet, ensure that unauthorized writes are blocked, duplicate its contents to a separate editable sheet, and confirm that the copy can be modified.
// Use Cases: Create a read‑only template sheet while providing a separate editable sheet for user input. | Generate modifiable reports from a protected source workbook without exposing protection settings. | Automated testing of worksheet protection to guarantee data integrity before distribution.
// AI Prompts: Write C# code using Aspose.Cells that protects a worksheet with a password, catches the exception on an illegal write, copies the sheet to a new unprotected worksheet, and verifies editability. | Explain how Worksheet.Copy handles protection flags in Aspose.Cells and show how to programmatically check the IsProtected property on source and destination sheets. | Provide a step‑by‑step guide to test worksheet protection: protect, attempt write, copy, confirm IsProtected status, edit copied sheet, and save the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    // This C# example demonstrates how to protect a worksheet with a password using Aspose.Cells, confirm that write attempts raise an exception, copy the protected sheet's content to a new unprotected worksheet, check the IsProtected flag on both sheets, edit the copied sheet successfully, and save the workbook.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet protectedSheet = workbook.Worksheets[0];
            Cells cells = protectedSheet.Cells;

            // Fill some data into the protected sheet
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue(123);
            cells["B2"].PutValue(456);

            // Protect the worksheet with a password and all protection types
            protectedSheet.Protect(ProtectionType.All, "myPassword", null);

            // Verify that the worksheet is protected
            Console.WriteLine($"Worksheet \"{protectedSheet.Name}\" IsProtected: {protectedSheet.IsProtected}");

            // Attempt to modify a cell in the protected worksheet (should throw an exception)
            try
            {
                cells["A3"].PutValue("Attempted Write");
                Console.WriteLine("Write to protected sheet succeeded (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing to protected sheet: {ex.Message}");
            }

            // Add a new worksheet that will remain unprotected
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet unprotectedSheet = workbook.Worksheets[newSheetIndex];

            // Copy the contents of the protected sheet to the unprotected sheet
            // (Worksheet.Copy copies cells, formats, etc., regardless of protection)
            protectedSheet.Copy(unprotectedSheet);

            // Verify that the target sheet is not protected
            Console.WriteLine($"Worksheet \"{unprotectedSheet.Name}\" IsProtected: {unprotectedSheet.IsProtected}");

            // Attempt to modify a cell in the unprotected worksheet (should succeed)
            try
            {
                unprotectedSheet.Cells["A3"].PutValue("Write Successful");
                Console.WriteLine("Write to unprotected sheet succeeded.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing to unprotected sheet: {ex.Message}");
            }

            // Save the workbook to verify the results
            workbook.Save("ProtectedAndCopied.xlsx");
        }
    }
}
