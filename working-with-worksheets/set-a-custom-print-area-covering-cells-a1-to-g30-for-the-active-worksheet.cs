// Title: Set a custom print area (A1:G30) on the active worksheet using Aspose.Cells for C#
// AI Prompts: Generate C# code that creates a workbook, selects the first worksheet, and assigns the printable range A1:G30 using the worksheet settings API, then saves the file. | Show how to use Aspose.Cells to define a custom printable region for a worksheet and persist the workbook. | Provide a concise example that demonstrates configuring the worksheet's print configuration to limit printing to cells A1 through G30 in C#.
// Common Searches: Aspose.Cells C# how to set print area for a worksheet | example code to define print range A1 to G30 in Aspose.Cells | C# Aspose.Cells set PageSetup.PrintArea property and save workbook | programmatically set Excel print area using Aspose.Cells library | set custom print area for first worksheet Aspose.Cells C#
// Tags: Aspose.Cells PageSetup.PrintArea | C# set worksheet printable range | Aspose.Cells export workbook with custom print region | Excel printable area configuration using Aspose.Cells | define custom print range in C# Aspose.Cells

using Aspose.Cells;

// // This program creates a new workbook, accesses the first worksheet, sets its printable range to A1:G30, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Get the active worksheet (first worksheet by default)
        Worksheet worksheet = workbook.Worksheets[0];

        // Set a custom print area covering cells A1 to G30
        worksheet.PageSetup.PrintArea = "A1:G30";

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
