// Title: Set the RefersTo property of a named range to a SUM formula with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set RefersTo of a named range to a formula | create a named formula that sums a range using Aspose.Cells .NET | assign a SUM expression to a named range programmatically in C# | how to use workbook.Names.Add and RefersTo to define a total sum in Aspose.Cells | reference worksheet cells in a named formula with Aspose.Cells C#
// Tags: RefersTo property assignment Aspose.Cells | named formula creation SUM Aspose.Cells | adding named range workbook.Names C# | using named formula in worksheet Aspose.Cells | saving workbook to XLSX Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, adds values to A1:A3, defines a named formula "=SUM(Data!A1:A3)", adds a named range called "TotalSum", sets its RefersTo property to the formula, uses the name in cell B1, and saves the file as NamedFormulaExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // Add sample data for the calculation
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);

        // Define a named formula that sums the range A1:A3
        // The formula uses the worksheet name to be fully qualified
        string formula = "=SUM(Data!A1:A3)";

        // Add a named range (named formula) to the workbook
        // The name "TotalSum" will refer to the formula defined above
        int nameIndex = workbook.Worksheets.Names.Add("TotalSum");
        Name totalSumName = workbook.Worksheets.Names[nameIndex];
        totalSumName.RefersTo = formula; // Set the RefersTo property

        // Use the named formula in a cell (optional)
        sheet.Cells["B1"].Formula = "TotalSum";

        // Save the workbook to a file
        workbook.Save("NamedFormulaExample.xlsx");
    }
}
