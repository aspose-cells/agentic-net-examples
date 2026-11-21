// Title: How to repeat columns A and B as print titles on every page with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set columns A:B as print titles for each printed page | repeat first two columns on every printed sheet using Aspose.Cells PageSetup | C# code to define print title columns in an Excel workbook with Aspose.Cells | example of PageSetup.PrintTitleColumns for .xlsx files in Aspose.Cells | programmatically set print title columns in an Aspose.Cells workbook
// Tags: Aspose.Cells PageSetup PrintTitleColumns | C# repeat columns on printed pages | Excel workbook print titles Aspose.Cells | set print title columns .NET | Aspose.Cells header columns printing

using Aspose.Cells;

// Demonstrates creating a workbook, configuring the first worksheet to repeat columns A and B on each printed page via PageSetup.PrintTitleColumns, adding sample header and data cells, and saving the file as Output.xlsx.
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

        // (Optional) Add sample data to demonstrate the titles
        sheet.Cells["A1"].PutValue("Header A");
        sheet.Cells["B1"].PutValue("Header B");
        sheet.Cells["A2"].PutValue("Data A1");
        sheet.Cells["B2"].PutValue("Data B1");

        // Save the workbook to a file
        workbook.Save("Output.xlsx", SaveFormat.Xlsx);
    }
}
