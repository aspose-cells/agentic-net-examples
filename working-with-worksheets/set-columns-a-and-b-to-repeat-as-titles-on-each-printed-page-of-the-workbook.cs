// Title: Set columns A and B as repeating print titles on each printed page using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells PageSetup.PrintTitleColumns to make columns A and B repeat as titles on every printed worksheet page in C#. | Write C# code that configures a workbook to repeat specific columns as print titles and then saves the file with Aspose.Cells. | Adjust the example to repeat columns A through C as print titles on each printed page using Aspose.Cells.
// Common Searches: Aspose.Cells C# repeat columns A and B on every printed page | How to set print title columns in an Excel workbook with Aspose.Cells for .NET | PageSetup.PrintTitleColumns example in C# Aspose.Cells | Set repeating column headers for printing using Aspose.Cells .NET | Configure print titles for columns in Excel via Aspose.Cells C# code
// Tags: Aspose.Cells PageSetup.PrintTitleColumns | C# repeat columns as print titles | Excel worksheet print title columns Aspose | configure worksheet print titles .NET | Aspose.Cells repeat column headers on print

using Aspose.Cells;

// // Creates a new workbook, accesses the first worksheet, sets columns A and B to repeat as print titles on every printed page via PageSetup.PrintTitleColumns, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set columns A and B to repeat as titles on each printed page
        sheet.PageSetup.PrintTitleColumns = "$A:$B";

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
