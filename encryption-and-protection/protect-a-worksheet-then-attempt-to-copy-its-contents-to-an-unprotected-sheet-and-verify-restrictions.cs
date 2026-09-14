// Title: How to protect a worksheet with a password, copy it to another sheet, and ensure the copy is unprotected using Aspose.Cells for .NET
// AI Prompts: Generate C# code that locks a worksheet with a password, attempts a prohibited edit, copies the worksheet to a new sheet, clears its protection, and confirms that the copied sheet can be edited using Aspose.Cells. | Demonstrate catching the error raised when trying to modify a locked cell on the original sheet and then successfully updating the same cell on the unprotected copied sheet in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells .NET protect worksheet with password and copy to new sheet unprotected | C# copy protected worksheet and remove protection after copy using Aspose.Cells | How to handle exception when editing a protected cell in Aspose.Cells | Verify worksheet protection status after copying in Aspose.Cells C# | Save workbook after unprotecting copied worksheet Aspose.Cells
// Tags: password-protect sheet Aspose.Cells .NET | copy sheet and clear protection Aspose.Cells | unprotect copied sheet Aspose.Cells | handle edit exception on protected sheet Aspose.Cells | save workbook after sheet protection changes Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, applies password protection to the first worksheet, catches the exception from an illegal edit, copies the sheet to a new worksheet, removes protection from the copy, verifies that edits are now allowed, and saves the workbook.
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
            sourceSheet.Name = "ProtectedSheet";
            sourceSheet.Cells["A1"].PutValue("Header");
            sourceSheet.Cells["A2"].PutValue("Data1");
            sourceSheet.Cells["A3"].PutValue("Data2");

            // Protect the worksheet with a password (oldPassword is not required, pass null)
            sourceSheet.Protect(ProtectionType.All, "Secret123", null);

            // Attempt to modify a cell in the protected sheet (should fail)
            try
            {
                sourceSheet.Cells["A2"].PutValue("ModifiedData");
                Console.WriteLine("Unexpected: Modification succeeded on a protected sheet.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Expected exception when modifying protected sheet: " + ex.Message);
            }

            // Add a new unprotected worksheet
            int destIndex = workbook.Worksheets.Add();
            Worksheet destSheet = workbook.Worksheets[destIndex];
            destSheet.Name = "UnprotectedCopy";

            // Copy contents from the protected sheet to the new sheet
            // This copies cells, formats, etc., but also copies protection settings
            sourceSheet.Copy(destSheet);

            // Ensure the destination sheet is unprotected
            destSheet.Unprotect();

            // Verify that the destination sheet is not protected
            if (!destSheet.IsProtected)
            {
                Console.WriteLine("Destination sheet is unprotected as expected.");
            }
            else
            {
                Console.WriteLine("Unexpected: Destination sheet is still protected.");
            }

            // Attempt to modify a cell in the copied (unprotected) sheet (should succeed)
            try
            {
                destSheet.Cells["A2"].PutValue("ModifiedInCopy");
                Console.WriteLine("Modification succeeded on the unprotected copied sheet.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected exception when modifying unprotected sheet: " + ex.Message);
            }

            // Save the workbook to verify the result
            string outputPath = "WorksheetProtectionDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error saving workbook: " + ex.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
