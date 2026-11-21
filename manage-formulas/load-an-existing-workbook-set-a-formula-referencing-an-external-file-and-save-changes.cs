// Title: Set an external workbook reference formula in an existing Excel file with Aspose.Cells for .NET (C#)
// Common Searches: how to add a formula that links to another Excel workbook using Aspose.Cells in C# | Aspose.Cells C# set cell formula referencing external file and save workbook | example of external workbook reference formula with Aspose.Cells .NET | update existing Excel file with external link formula using Aspose.Cells API
// Tags: set external workbook formula Aspose.Cells | update cell formula C# Aspose.Cells | reference external Excel file formula .NET | save workbook after formula modification Aspose.Cells | load workbook assign external link formula C#

using Aspose.Cells;

// The sample loads an existing workbook, assigns cell A1 a formula that references B2 in Sheet1 of an external Excel file, and saves the updated workbook to a new file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Access the first worksheet (you can change the index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Set a formula in cell A1 that references an external workbook
        // Example: =('[ExternalFile.xlsx]Sheet1'!B2)
        sheet.Cells["A1"].Formula = "='[ExternalFile.xlsx]Sheet1'!B2";

        // Save the workbook with the updated formula
        workbook.Save("OutputWorkbook.xlsx");
    }
}
