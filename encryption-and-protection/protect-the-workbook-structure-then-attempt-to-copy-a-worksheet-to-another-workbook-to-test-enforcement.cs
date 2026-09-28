// Title: C# example: protect workbook structure with Aspose.Cells and verify that copying a protected worksheet to another workbook is blocked
// AI Prompts: Show how to protect a workbook's structure using Aspose.Cells in C# and then attempt to copy its worksheet to a new workbook, handling the expected CellsException. | Write C# code that creates a source workbook, applies structure protection without a password, saves it, tries Worksheets.AddCopy on a destination workbook, and captures the exception when the copy is disallowed.
// Common Searches: Aspose.Cells C# prevent copying sheet from structure‑protected workbook | How to catch CellsException when adding a copy of a protected worksheet in .NET | Copy worksheet from password‑less protected Excel file using Aspose.Cells throws error | Aspose.Cells protect workbook structure without password and test copy restriction
// Tags: Aspose.Cells protect workbook structure | Aspose.Cells copy worksheet from protected workbook | C# Worksheets.AddCopy protected workbook exception | handle CellsException when copying sheet | structure protection enforcement Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a source workbook, adds data, protects its structure with an empty password, saves it, then creates a destination workbook and attempts to copy the protected worksheet using Worksheets.AddCopy. The operation throws a CellsException, which is caught and logged, demonstrating that structure protection blocks worksheet copying.
class WorkbookStructureProtectionDemo
{
    static void Main()
    {
        try
        {
            // Create the source workbook and add some data
            Workbook sourceWorkbook = new Workbook();
            Worksheet sheet = sourceWorkbook.Worksheets[0];
            sheet.Name = "SourceSheet";
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B2"].PutValue(123);

            // Protect the workbook structure (empty password)
            sourceWorkbook.Protect(ProtectionType.Structure, string.Empty);

            // Save the source workbook (optional, just to demonstrate persistence)
            string sourcePath = "SourceProtected.xlsx";
            sourceWorkbook.Save(sourcePath);

            // Create a destination workbook
            Workbook destWorkbook = new Workbook();

            try
            {
                // Attempt to copy the protected worksheet to the destination workbook
                // This should fail because the source workbook's structure is protected
                // Use the overload that copies a worksheet by its name
                destWorkbook.Worksheets.AddCopy(sheet.Name);
                Console.WriteLine("Worksheet copied successfully (unexpected).");
            }
            catch (CellsException ex)
            {
                // Expected exception when trying to copy from a protected workbook
                Console.WriteLine("Copy operation failed as expected: " + ex.Message);
            }

            // Save the destination workbook (will be empty if copy failed)
            string destPath = "Destination.xlsx";
            destWorkbook.Save(destPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}
