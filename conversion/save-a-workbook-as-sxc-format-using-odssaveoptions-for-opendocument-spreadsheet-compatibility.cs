// Title: How to save a workbook as SXC (OpenDocument Spreadsheet) with Aspose.Cells using OdsSaveOptions in C#
// Common Searches: C# Aspose.Cells export workbook to SXC file | How to use OdsSaveOptions for SXC output in .NET | Convert Excel workbook to OpenDocument Spreadsheet (SXC) with Aspose.Cells | Saving workbook as SXC using SaveFormat.SXC in C# | Set ODS options when generating SXC with Aspose.Cells
// Tags: OdsSaveOptions for SXC generation | Aspose.Cells OpenDocument Spreadsheet export | C# workbook save to SXC format | SaveFormat.SXC with Aspose.Cells | OpenDocument spreadsheet creation using .NET

using Aspose.Cells;
using Aspose.Cells.Saving;

// The example creates a new Workbook, adds data, configures OdsSaveOptions with SaveFormat.SXC, and saves the workbook as an SXC file compatible with OpenDocument Spreadsheet.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample");
        sheet.Cells["B1"].PutValue("Data");

        // Configure ODS save options to produce SXC (OpenDocument Spreadsheet) format
        OdsSaveOptions saveOptions = new OdsSaveOptions(SaveFormat.SXC);
        // Additional ODS options can be set here if needed, e.g., saveOptions.CheckWorksheetNames = true;

        // Save the workbook as an SXC file using the configured options
        workbook.Save("SampleOutput.sxc", saveOptions);
    }
}
