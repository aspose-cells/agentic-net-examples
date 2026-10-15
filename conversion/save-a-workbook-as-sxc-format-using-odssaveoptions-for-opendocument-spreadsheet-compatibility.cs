// Title: How to export an Excel workbook to SXC (OpenDocument Spreadsheet) using OdsSaveOptions in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# example for saving a workbook as .sxc | using OdsSaveOptions to export Excel to OpenDocument Spreadsheet format | programmatically create SXC file with Aspose.Cells .NET API | convert Excel workbook to SXC with SaveFormat.SXC in C# | save workbook as OpenDocument Spreadsheet using Aspose.Cells OdsSaveOptions | C# code to generate .sxc file from Aspose.Cells workbook
// Tags: OdsSaveOptions SaveFormat.SXC export | Aspose.Cells OpenDocument Spreadsheet generation | C# Excel to SXC conversion | Aspose.Cells .NET create .sxc file | Export workbook to SXC with Aspose.Cells | OpenDocument Spreadsheet output using Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a Workbook, adds data, configures OdsSaveOptions with SaveFormat.SXC, and saves the file as "output.sxc", providing OpenDocument Spreadsheet compatibility via Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample");
        sheet.Cells["B1"].PutValue(123);
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Configure ODS save options to output SXC (OpenDocument Spreadsheet) format
        OdsSaveOptions saveOptions = new OdsSaveOptions(SaveFormat.SXC);

        // Save the workbook as SXC using the specified options
        workbook.Save("output.sxc", saveOptions);
    }
}
