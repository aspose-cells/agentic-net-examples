// Title: Set columns A and B as repeating print titles on each printed page with Aspose.Cells for .NET (C#)
// AI Prompts: In C#, use Aspose.Cells to assign "$A:$B" to Worksheet.PageSetup.PrintTitleColumns and save the workbook as an XLSX file. | Create a workbook, access the first worksheet, configure it to repeat the first two columns on every printed page, then write the file using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to repeat first two columns on printed pages | set print title columns A B using Aspose.Cells .NET | C# Aspose.Cells PageSetup PrintTitleColumns example code | repeat column headers on each sheet printout Aspose.Cells | configure worksheet print titles for columns in Aspose.Cells C#
// Tags: Aspose.Cells PrintTitleColumns property | Aspose.Cells column header repetition | C# worksheet page setup print titles | Aspose.Cells set repeated columns A B | export workbook with print title columns Aspose.Cells

using Aspose.Cells;

// The program creates a new workbook, accesses the first worksheet, sets the PrintTitleColumns property to "$A:$B" so columns A and B repeat on every printed page, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the print title columns to columns A and B (repeated column headers)
        sheet.PageSetup.PrintTitleColumns = "$A:$B";

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
