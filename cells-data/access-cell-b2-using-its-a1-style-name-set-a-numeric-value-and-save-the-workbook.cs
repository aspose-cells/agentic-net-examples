// Title: Write a numeric value to cell B2 using A1 notation with Aspose.Cells for .NET and save as an XLSX file
// Common Searches: Aspose.Cells C# set numeric value in cell B2 using A1 address | How to write data to a specific cell with A1 style reference in Aspose.Cells .NET | Saving a workbook after updating a single cell with Aspose.Cells for .NET | C# example for accessing cell B2 via A1 notation in Aspose.Cells | Create and save an Excel file after modifying cell value using Aspose.Cells
// Tags: set cell value using A1 notation Aspose.Cells | write numeric data to worksheet cell C# | save workbook as xlsx Aspose.Cells | access cell by address Aspose.Cells .NET | modify specific cell in Excel file using Aspose.Cells | Aspose.Cells cell PutValue example

using Aspose.Cells;

// The sample creates a new Workbook, retrieves the first worksheet, accesses cell B2 via its A1 address, assigns the numeric value 123.45 using PutValue, and saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Access cell B2 using its A1 style name and set a numeric value
        Cell cell = sheet.Cells["B2"];
        cell.PutValue(123.45);

        // Save the workbook to a file
        workbook.Save("Result.xlsx");
    }
}
