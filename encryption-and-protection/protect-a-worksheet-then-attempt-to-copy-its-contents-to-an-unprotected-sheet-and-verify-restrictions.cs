// Title: C# – Protect a Worksheet with a Password, Copy Its Data to an Unprotected Sheet, and Verify Protection Is Not Transferred (Aspose.Cells)
// Description: Demonstrates how to protect the first worksheet of a workbook using Aspose.Cells for .NET, copy only its cell values and formatting to a new unprotected worksheet in another workbook, confirm that the protection flag (IsProtected) remains unchanged, and save both workbooks.
// Keywords: Aspose.Cells protect worksheet | worksheet password .NET | Worksheet.Copy protection | IsProtected property | copy sheet without protection | C# Aspose.Cells example | save protected workbook | unprotected destination sheet
// Common Searches: How to protect a worksheet with a password using Aspose.Cells C# | Copy data from a protected sheet to another workbook without copying protection | Does Worksheet.Copy retain protection settings in Aspose.Cells? | Check IsProtected after copying a worksheet | Save protected and unprotected workbooks with Aspose.Cells
// Developer Intent: Protect a worksheet, duplicate its contents to an unprotected sheet, and ensure the protection settings are not carried over.
// Use Cases: Create a read‑only template, protect it, then generate an editable copy for end‑users. | Extract data from a secured source workbook and place it into a new workbook for analysis without retaining protection. | Automated testing to verify that worksheet protection does not propagate when sheets are copied across workbooks.
// AI Prompts: Generate C# code with Aspose.Cells that protects a worksheet using a password, copies only cell values and formatting to another workbook, and confirms the destination sheet stays unprotected. | Explain the behavior of the IsProtected property when Worksheet.Copy is invoked in Aspose.Cells. | Provide a step‑by‑step tutorial to protect a sheet, copy its contents to an unprotected sheet, and validate protection status before and after the copy.

using System;
using Aspose.Cells;

namespace AsposeCellsProtectionCopyDemo
{
    // Demonstrates how to protect the first worksheet of a workbook using Aspose.Cells for .NET, copy only its cell values and formatting to a new unprotected worksheet in another workbook, confirm that the protection flag (IsProtected) remains unchanged, and save both workbooks.
    public class Program
    {
        public static void Main()
        {
            // ---------- Create source workbook and protect its first worksheet ----------
            Workbook sourceWorkbook = new Workbook();                     // create workbook
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];        // get first worksheet

            // Populate some data
            sourceSheet.Cells["A1"].PutValue("Protected Sheet");
            sourceSheet.Cells["B2"].PutValue(12345);

            // Protect the worksheet with all protection types and a password
            sourceSheet.Protect(ProtectionType.All, "pwd123", null);

            // Verify protection status
            Console.WriteLine($"Source sheet IsProtected: {sourceSheet.IsProtected}"); // should be True

            // ---------- Create destination workbook (unprotected) ----------
            Workbook destWorkbook = new Workbook();                       // create another workbook
            Worksheet destSheet = destWorkbook.Worksheets[0];            // get its first worksheet

            // Ensure destination sheet is not protected
            Console.WriteLine($"Destination sheet initially IsProtected: {destSheet.IsProtected}"); // should be False

            // ---------- Attempt to copy contents from protected source to unprotected destination ----------
            // This copies cells and formats but does NOT copy protection settings.
            destSheet.Copy(sourceSheet);

            // Verify that the copy succeeded (cell values are transferred)
            Console.WriteLine($"Copied cell A1 value: {destSheet.Cells["A1"].StringValue}");
            Console.WriteLine($"Copied cell B2 value: {destSheet.Cells["B2"].IntValue}");

            // Verify protection status after copy
            Console.WriteLine($"Source sheet still IsProtected: {sourceSheet.IsProtected}"); // should remain True
            Console.WriteLine($"Destination sheet after copy IsProtected: {destSheet.IsProtected}"); // should be False

            // ---------- Save both workbooks ----------
            sourceWorkbook.Save("ProtectedSource.xlsx");
            destWorkbook.Save("UnprotectedDestination.xlsx");
        }
    }
}
