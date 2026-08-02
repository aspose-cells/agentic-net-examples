using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – preserve relative formula references when copying rows
class PreserveFormulaCopyDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sheet1";

        // Populate some sample data
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B1"].PutValue(20);
        // Set a formula that uses relative references (A1+B1)
        sheet.Cells["C1"].Formula = "=A1+B1";

        // Insert a blank row at index 1 (Excel row 2)
        sheet.Cells.InsertRows(1, 1);

        // Configure copy options to keep relative references correct
        CopyOptions copyOptions = new CopyOptions
        {
            // If the source sheet has been copied before, keep references to sheets with the same name
            ReferToSheetWithSameName = true,
            // Extend copying to adjacent ranges (e.g., merged cells) if present
            ExtendToAdjacentRange = true
        };

        // Copy row 0 (first row) to row 1 (second row) using the configured options
        sheet.Cells.CopyRows(sheet.Cells, 0, 1, 1, copyOptions);

        // Output the formulas to verify that the copied formula has been updated to reference the new row
        Console.WriteLine("Original formula (C1): " + sheet.Cells["C1"].Formula); // Expected: =A1+B1
        Console.WriteLine("Copied formula   (C2): " + sheet.Cells["C2"].Formula); // Expected: =A2+B2

        // Save the workbook
        workbook.Save("PreserveFormulaCopy.xlsx");
    }
}