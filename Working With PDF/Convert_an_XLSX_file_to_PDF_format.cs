using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Source Excel file (XLSX)
        string sourcePath = "input.xlsx";

        // Destination PDF file
        string destPath = "output.pdf";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Save the workbook as PDF
        workbook.Save(destPath, SaveFormat.Pdf);

        Console.WriteLine("Conversion completed successfully.");
    }
}