using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Example: add some data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data");
        sheet.Cells["B2"].PutValue(12345);

        // Disable any automatic scaling for every worksheet
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Use percent scaling (Zoom) instead of FitToPages
            ws.PageSetup.IsPercentScale = true;
            // Set Zoom to 100% to keep original dimensions exactly
            ws.PageSetup.Zoom = 100;
        }

        // Save the workbook as PDF preserving the original layout
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}