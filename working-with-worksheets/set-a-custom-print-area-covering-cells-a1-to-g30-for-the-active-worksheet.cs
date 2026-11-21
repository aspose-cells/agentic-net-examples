// Title: Set a custom print area (A1:G30) on the active worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set print area for specific cell range | how to define custom print area A1:G30 in a .NET workbook | C# Aspose.Cells PageSetup PrintArea example code | set active worksheet print range using Aspose.Cells API | configure Excel print area programmatically with Aspose.Cells
// Tags: Aspose.Cells worksheet print area configuration | C# PageSetup PrintArea usage | custom Excel print range A1:G30 in .NET | active worksheet print settings Aspose.Cells | save workbook after print area setup | Aspose.Cells API for print area definition

using Aspose.Cells;

// The example creates a new Workbook, accesses the first (active) worksheet, assigns the PageSetup.PrintArea property to "A1:G30" to establish a custom print area, and then saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the active worksheet (first worksheet by default)
        Worksheet sheet = workbook.Worksheets[0];

        // Set the custom print area to cover cells A1 through G30
        sheet.PageSetup.PrintArea = "A1:G30";

        // Save the workbook to verify the setting (optional)
        workbook.Save("output.xlsx");
    }
}
