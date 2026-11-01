// Title: How to convert an Excel workbook to a high‑quality XPS document using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# save workbook as XPS for printing | convert .xlsx file to XPS format with Aspose.Cells API | C# code example for exporting Excel to XPS using SaveFormat.Xps | generate high resolution XPS from Excel workbook in .NET | Aspose.Cells XPS output settings for print quality
// Tags: Aspose.Cells SaveFormat.Xps usage | C# Excel to XPS conversion | XPS generation from workbook for printing | .NET high‑resolution XPS export | Aspose.Cells workbook save as XPS

using System;
using Aspose.Cells;

// The sample loads an Excel file (input.xlsx) into an Aspose.Cells Workbook and saves it as an XPS document (output.xps) by specifying SaveFormat.Xps, delivering a high‑quality XPS file suitable for printing.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Convert and save the workbook as an XPS document for high‑quality printing
        workbook.Save("output.xps", SaveFormat.Xps);
    }
}
