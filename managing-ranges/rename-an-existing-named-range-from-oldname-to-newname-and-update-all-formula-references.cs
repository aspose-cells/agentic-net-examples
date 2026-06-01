using System;
using Aspose.Cells;

class RenameNamedRange
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sheet1";

        // Populate some sample data
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);

        // Create a named range called "OldName" that refers to A1:A3
        int nameIndex = workbook.Worksheets.Names.Add("OldName");
        Name oldName = workbook.Worksheets.Names[nameIndex];
        oldName.RefersTo = "=Sheet1!$A$1:$A$3";

        // Use the named range in a formula
        sheet.Cells["B1"].Formula = "=SUM(OldName)";

        // Calculate the formula before renaming
        workbook.CalculateFormula();
        Console.WriteLine("Sum before rename: " + sheet.Cells["B1"].Value); // Expected: 60

        // Rename the named range to "NewName"
        oldName.Text = "NewName";

        // Recalculate after renaming; Aspose.Cells updates formula references automatically
        workbook.CalculateFormula();

        // Verify that the formula now references the new name
        Console.WriteLine("Formula after rename: " + sheet.Cells["B1"].Formula); // Expected: =SUM(NewName)
        Console.WriteLine("Sum after rename: " + sheet.Cells["B1"].Value); // Expected: 60

        // Save the workbook
        workbook.Save("RenamedNamedRange.xlsx");
    }
}