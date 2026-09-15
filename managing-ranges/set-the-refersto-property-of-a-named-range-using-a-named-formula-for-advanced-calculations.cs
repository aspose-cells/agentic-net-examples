// Title: How to set a SUM formula in the RefersTo property of a named range using Aspose.Cells for .NET (C#)
// AI Prompts: Create a named range called MyRange and assign its RefersTo property to "=SUM(Data!A1:A5)" using Aspose.Cells in C#. | Update an existing named range to reference a worksheet formula via the RefersTo property with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set RefersTo of a named range to a SUM formula | programmatically assign a formula to a named range in a .NET workbook | how to define a named range that calculates total of a column using Aspose.Cells | C# example for using RefersTo property with worksheet formula in Aspose.Cells
// Tags: set RefersTo property named range Aspose.Cells | assign SUM formula to named range C# | named range formula definition Aspose.Cells | Aspose.Cells workbook create named range with formula | C# RefersTo named range calculation

using Aspose.Cells;

// The sample creates a workbook, adds a worksheet named "Data", fills cells A1‑A5 with numbers, defines a named range "MyRange", sets its RefersTo property to the formula "=SUM(Data!A1:A5)", and saves the file as "AdvancedNamedRange.xlsx" using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Add a worksheet to hold data
        int sheetIdx = workbook.Worksheets.Add();
        Worksheet sheet = workbook.Worksheets[sheetIdx];
        sheet.Name = "Data";

        // Populate some sample values (optional, just for demonstration)
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);
        sheet.Cells["A4"].PutValue(40);
        sheet.Cells["A5"].PutValue(50);

        // Create a named range called "MyRange"
        int nameIdx = workbook.Worksheets.Names.Add("MyRange");

        // Set the RefersTo property using a named formula.
        // This formula will calculate the sum of the range A1:A5 on the "Data" sheet.
        workbook.Worksheets.Names[nameIdx].RefersTo = "=SUM(Data!A1:A5)";

        // Save the workbook (lifecycle rule)
        workbook.Save("AdvancedNamedRange.xlsx");
    }
}
