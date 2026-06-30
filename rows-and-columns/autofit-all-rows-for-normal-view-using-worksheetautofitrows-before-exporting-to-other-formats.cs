using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample data
        worksheet.Cells["A1"].PutValue("Sample text for AutoFitRows demonstration");
        worksheet.Cells["A2"].PutValue("Another line\nwith line break to test row height adjustment");

        // Auto‑fit all rows in normal view
        worksheet.AutoFitRows();

        // Export to PDF
        string outputPath = "AutoFitRowsDemo.pdf";
        workbook.Save(outputPath, SaveFormat.Pdf);
    }
}

// Author: Example demonstrating AutoFitRows before export using Aspose.Cells.