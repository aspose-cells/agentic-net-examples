// Title: Aspose.Cells .NET – Protect Workbook Structure and Verify Copy Restrictions
// Description: Demonstrates how to protect a workbook's structure with a password, catch the exception thrown by Worksheets.AddCopy on the protected file, and successfully copy the sheet to a separate workbook.
// Keywords: Aspose.Cells protect structure | C# workbook protection | AddCopy exception | copy worksheet to another workbook | Excel file security Aspose | structure protection example
// Common Searches: protect workbook structure Aspose.Cells C# | AddCopy fails after protecting workbook | copy sheet from protected workbook Aspose | Aspose.Cells worksheet copy restrictions
// Developer Intent: Show that enabling structure protection blocks internal sheet duplication while still allowing the sheet to be copied to a different workbook.
// Use Cases: Enforce read‑only layout by preventing users from adding, deleting, or moving worksheets in the original file. | Validate protection settings by handling the exception from Worksheets.AddCopy on a protected workbook. | Export a protected sheet to a new workbook for reporting or distribution without altering the source.
// AI Prompts: Write C# code using Aspose.Cells to protect a workbook's structure with a password, attempt Worksheets.AddCopy, and handle the expected exception. | Explain how to copy a worksheet from a password‑protected workbook to another workbook while respecting structure protection in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to protect a workbook's structure with a password, catch the exception thrown by Worksheets.AddCopy on the protected file, and successfully copy the sheet to a separate workbook.
class WorkbookStructureProtectionDemo
{
    static void Main()
    {
        // -------------------------------------------------
        // Create a source workbook and put some sample data
        // -------------------------------------------------
        Workbook sourceWorkbook = new Workbook();
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
        sourceSheet.Name = "SampleSheet";
        sourceSheet.Cells["A1"].PutValue("Protected Workbook");
        sourceSheet.Cells["A2"].PutValue(42);

        // -------------------------------------------------
        // Protect the workbook structure with a password
        // -------------------------------------------------
        sourceWorkbook.Protect(ProtectionType.Structure, "pwd123");

        // Save the protected workbook (optional, just to see the file)
        sourceWorkbook.Save("ProtectedSource.xlsx");

        // -------------------------------------------------
        // Attempt to copy a worksheet inside the protected workbook
        // This operation should be blocked because the structure is protected
        // -------------------------------------------------
        try
        {
            // AddCopy tries to add a new worksheet based on an existing one
            sourceWorkbook.Worksheets.AddCopy(0);
            Console.WriteLine("AddCopy succeeded unexpectedly.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("AddCopy failed as expected: " + ex.Message);
        }

        // -------------------------------------------------
        // Create a separate destination workbook
        // -------------------------------------------------
        Workbook destinationWorkbook = new Workbook();

        // -------------------------------------------------
        // Copy the worksheet from the protected source workbook
        // to the destination workbook – this is allowed because
        // we are not modifying the protected workbook's structure
        // -------------------------------------------------
        try
        {
            Worksheet destSheet = destinationWorkbook.Worksheets[0];
            destSheet.Name = "CopiedSheet";
            destSheet.Copy(sourceSheet);
            Console.WriteLine("Worksheet copied to destination workbook successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Worksheet copy to destination failed: " + ex.Message);
        }

        // Save the destination workbook to verify the result
        destinationWorkbook.Save("Destination.xlsx");
    }
}
