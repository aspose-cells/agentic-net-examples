using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the HTML file into a workbook.
        // The Workbook constructor automatically detects the format from the file extension.
        Workbook workbook = new Workbook("input.html");

        // Save the workbook as an XLSX file.
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}