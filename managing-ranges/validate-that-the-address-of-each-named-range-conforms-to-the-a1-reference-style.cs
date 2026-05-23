using Aspose.Cells;
using System;

class ValidateNamedRangesA1
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Name = "Sheet1";

        // Populate some sample data
        ws.Cells["A1"].PutValue(1);
        ws.Cells["A2"].PutValue(2);
        ws.Cells["A3"].PutValue(3);
        ws.Cells["B1"].PutValue(10);
        ws.Cells["B2"].PutValue(20);
        ws.Cells["B3"].PutValue(30);

        // Create a named range using A1 reference style
        int idxA1 = wb.Worksheets.Names.Add("A1Range");
        Name nameA1 = wb.Worksheets.Names[idxA1];
        nameA1.RefersTo = "=Sheet1!$A$1:$A$3";

        // Create a named range using R1C1 reference style
        int idxR1C1 = wb.Worksheets.Names.Add("R1C1Range");
        Name nameR1C1 = wb.Worksheets.Names[idxR1C1];
        nameR1C1.R1C1RefersTo = "'Sheet1'!R1C2:R3C2";

        // Validate each named range to ensure its address is in A1 style
        foreach (Name n in wb.Worksheets.Names)
        {
            // Get the reference formatted as A1
            string a1Reference = n.GetRefersTo(false, false);

            // Original stored reference (could be A1 or R1C1)
            string originalReference = n.RefersTo;

            // Determine if the original reference already matches A1 style
            bool conformsToA1 = string.Equals(originalReference, a1Reference, StringComparison.OrdinalIgnoreCase);

            Console.WriteLine($"Name: {n.Text}");
            Console.WriteLine($"Original RefersTo: {originalReference}");
            Console.WriteLine($"A1 RefersTo: {a1Reference}");
            Console.WriteLine($"Conforms to A1 style: {conformsToA1}");
            Console.WriteLine();
        }

        // Save the workbook
        wb.Save("ValidateNamedRangesA1.xlsx");
    }
}