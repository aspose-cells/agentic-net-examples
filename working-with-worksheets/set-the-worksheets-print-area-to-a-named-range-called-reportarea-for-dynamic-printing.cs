// Title: Set a worksheet's print area to the named range 'ReportArea' with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# assign a named range as the print area | set worksheet print area to a named range using PageSetup.PrintArea in .NET | dynamic print area configuration for Excel sheets with Aspose.Cells | C# code to set Excel print area to ReportArea named range | programmatically define print area in an Aspose.Cells workbook
// Tags: Aspose.Cells set print area named range | C# worksheet PageSetup PrintArea | Excel dynamic print area Aspose.Cells | programmatic print area configuration .NET | named range print area Aspose.Cells

using Aspose.Cells;

// The example loads an Excel workbook, accesses the first worksheet, assigns the PageSetup.PrintArea property to the named range "ReportArea", and saves the workbook, enabling dynamic printing based on the named range.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the worksheet where the print area should be set
        Worksheet worksheet = workbook.Worksheets[0]; // or workbook.Worksheets["Sheet1"]

        // Assign the print area to the named range "ReportArea"
        worksheet.PageSetup.PrintArea = "ReportArea";

        // Save the workbook (replace with your desired output file path)
        workbook.Save("output.xlsx");
    }
}
