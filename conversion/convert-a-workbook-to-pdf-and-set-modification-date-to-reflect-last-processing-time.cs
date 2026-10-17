// Title: Convert an Excel workbook to PDF and update its file modification timestamp using Aspose.Cells for .NET (C#)
// Common Searches: c# convert xlsx to pdf with Aspose.Cells and set file last modified date | how to change PDF file timestamp after saving from Excel using Aspose.Cells | Aspose.Cells save workbook as PDF then update file system timestamp in .NET | set last write time on generated PDF after Excel conversion in C# | update file modification time after Aspose.Cells PDF export
// Tags: Aspose.Cells workbook to PDF conversion C# | File.SetLastWriteTime usage after PDF generation | SaveFormat.Pdf parameter in Aspose.Cells | programmatically set file last write time .NET | Excel workbook PDF export timestamp handling

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook (input.xlsx), converts it to PDF (output.pdf) using Aspose.Cells' SaveFormat.Pdf, and then sets the PDF file's last write time to the current processing time with File.SetLastWriteTime.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Convert and save the workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);

        // Set the PDF file's modification date to the current processing time
        File.SetLastWriteTime("output.pdf", DateTime.Now);
    }
}
