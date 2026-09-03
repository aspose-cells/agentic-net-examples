// Title: Set a custom print area A1:G30 on the active worksheet with Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to set the PageSetup.PrintArea of the first worksheet to A1:G30 and then saves the workbook. | Show how to programmatically define a print range for a worksheet in Aspose.Cells and persist the changes to a .xlsx file.
// Common Searches: Aspose.Cells C# set print area for specific cells | How to define a print range A1:G30 in a .NET spreadsheet | Programmatically set worksheet print area using Aspose.Cells PageSetup | Saving a workbook after configuring print area with Aspose.Cells | C# example for custom print area in Aspose.Cells workbook
// Tags: Aspose.Cells PageSetup PrintArea property | C# set worksheet print range | custom print area A1:G30 Aspose.Cells | save workbook after print area configuration | active worksheet print area .NET

using Aspose.Cells;

// Creates a new Workbook, accesses the first worksheet, sets PageSetup.PrintArea to "A1:G30", and saves the file as PrintAreaDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (using the provided create rule)
        Workbook workbook = new Workbook();

        // Get the active worksheet (first worksheet by default)
        Worksheet sheet = workbook.Worksheets[0];

        // Set a custom print area covering cells A1 to G30
        sheet.PageSetup.PrintArea = "A1:G30";

        // Save the workbook (using the provided save rule)
        workbook.Save("PrintAreaDemo.xlsx");
    }
}
