using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example – load XLSX and save as PDF
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Path for the resulting PDF file
        string destPath = "output.pdf";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(sourcePath);

        // Save the loaded workbook as a PDF document
        workbook.Save(destPath, SaveFormat.Pdf);
    }
}