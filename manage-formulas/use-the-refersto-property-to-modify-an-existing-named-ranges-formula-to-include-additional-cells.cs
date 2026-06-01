using System;
using Aspose.Cells;

class ModifyNamedRange
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sheet1";

        // Populate some sample data
        sheet.Cells["A1"].PutValue(1);
        sheet.Cells["A2"].PutValue(2);
        sheet.Cells["A3"].PutValue(3);
        sheet.Cells["B1"].PutValue(4);
        sheet.Cells["B2"].PutValue(5);
        sheet.Cells["B3"].PutValue(6);

        // Add a named range that initially refers to A1:B2
        int nameIndex = workbook.Worksheets.Names.Add("MyRange");
        Name myRange = workbook.Worksheets.Names[nameIndex];
        myRange.RefersTo = "=Sheet1!$A$1:$B$2";

        // Modify the named range to include additional cells (extend to B3)
        myRange.RefersTo = "=Sheet1!$A$1:$B$3";

        // Use the updated named range in a formula
        sheet.Cells["C1"].Formula = "=SUM(MyRange)";
        workbook.CalculateFormula();

        // Output results
        Console.WriteLine("Sum of extended range: " + sheet.Cells["C1"].Value);
        Console.WriteLine("Updated RefersTo: " + myRange.RefersTo);

        // Save the workbook
        workbook.Save("ModifiedNamedRange.xlsx");
    }
}