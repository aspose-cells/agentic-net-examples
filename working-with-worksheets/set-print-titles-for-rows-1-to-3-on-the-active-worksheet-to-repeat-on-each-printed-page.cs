// Title: Set rows 1‑3 as repeat print titles on each printed page using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set rows 1 to 3 as print titles for each page | repeat header rows on every printed page with Aspose.Cells .NET | PageSetup.PrintTitleRows example in C# for Excel worksheets | how to configure print title rows in Aspose.Cells workbook | C# code to make first three rows repeat when printing an Excel file
// Tags: Aspose.Cells PageSetup.PrintTitleRows API | C# configure worksheet page setup | Excel print titles via Aspose.Cells | Set repeat rows in workbook using .NET | Aspose.Cells workbook print configuration

using Aspose.Cells;

// The example creates a workbook, accesses the first worksheet, assigns rows 1‑3 as print titles using the PageSetup.PrintTitleRows property, and saves the file as PrintTitlesExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Get the active worksheet (first worksheet by default)
        Worksheet sheet = workbook.Worksheets[0];

        // Set rows 1 to 3 as print titles so they repeat on each printed page
        // The range is specified in A1 style with absolute references
        sheet.PageSetup.PrintTitleRows = "$1:$3";

        // Save the workbook to verify the setting
        workbook.Save("PrintTitlesExample.xlsx");
    }
}
