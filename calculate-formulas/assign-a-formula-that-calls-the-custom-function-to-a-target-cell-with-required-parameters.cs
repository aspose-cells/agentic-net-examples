// Title: Assign a custom function formula with cell references to a worksheet cell using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set formula to call a user-defined function with cell arguments | How to use a custom Excel function in a workbook created with Aspose.Cells .NET | Assign MyCustomFunc(A1,B2) formula to a cell in Aspose.Cells C# example | Register and invoke custom functions in Aspose.Cells workbook programmatically | Set formula for cell C3 that calls a custom function in Aspose.Cells
// Tags: assign custom function formula Aspose.Cells | invoke user-defined Excel function C# | register custom functions collection Aspose.Cells | set cell formula with cell references .NET | Aspose.Cells workbook custom function usage

using Aspose.Cells;

// Creates a new workbook, writes values to A1 and B2, assigns the formula =MyCustomFunc(A1,B2) to cell C3, and saves the file. The custom function must be registered in the workbook's CustomFunctions collection beforehand.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set example values for the parameters of the custom function
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B2"].PutValue(20);

        // Assign a formula that calls the custom function MyCustomFunc with parameters A1 and B2
        // The custom function must be registered with the workbook's CustomFunctions collection elsewhere in the code.
        sheet.Cells["C3"].Formula = "=MyCustomFunc(A1,B2)";

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
