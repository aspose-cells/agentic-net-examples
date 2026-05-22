using System;
using Aspose.Cells;

class UnhideRowsAndExportPdf
{
    static void Main()
    {
        // Path to the existing Excel file
        string inputFile = "input.xlsx";

        // Path for the resulting PDF file
        string outputFile = "output.pdf";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Unhide rows 30 to 35 (zero‑based index). 
        // totalRows = 6 (rows 30,31,32,33,34,35). 
        // Height = -1 means auto‑fit (default height for hidden rows).
        worksheet.Cells.UnhideRows(30, 6, -1);

        // Save the workbook as PDF
        workbook.Save(outputFile, SaveFormat.Pdf);
    }
}