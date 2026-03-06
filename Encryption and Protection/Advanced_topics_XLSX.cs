using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class AdvancedXlsxDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Access the first worksheet and rename it
        Worksheet ws = wb.Worksheets[0];
        ws.Name = "Data";

        // Populate sample data
        ws.Cells["A1"].PutValue("Product");
        ws.Cells["B1"].PutValue("Quantity");
        ws.Cells["C1"].PutValue("Price");
        ws.Cells["A2"].PutValue("Apple");
        ws.Cells["B2"].PutValue(10);
        ws.Cells["C2"].PutValue(0.5);
        ws.Cells["A3"].PutValue("Banana");
        ws.Cells["B3"].PutValue(20);
        ws.Cells["C3"].PutValue(0.3);

        // Add a total row
        int totalRowIndex = 4; // Excel rows are 1‑based, so row 4 is index 3 (zero‑based)
        ws.Cells["A4"].PutValue("Total");
        ws.Cells["B4"].Formula = "SUM(B2:B3)";
        ws.Cells["C4"].Formula = "SUM(C2:C3)";

        // Set OOXML compliance to strict
        wb.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Save the workbook as XLSX
        string xlsxPath = "AdvancedDemo.xlsx";
        wb.Save(xlsxPath, SaveFormat.Xlsx);

        // Convert the XLSX file to PDF
        string pdfPath = "AdvancedDemo.pdf";
        ConversionUtility.Convert(xlsxPath, pdfPath);

        // Save the workbook with the total row as a separate file
        string outputWithSummary = "AdvancedDemo_WithSummary.xlsx";
        wb.Save(outputWithSummary, SaveFormat.Xlsx);

        // Inform the user
        Console.WriteLine("Demo completed. Generated files:");
        Console.WriteLine(xlsxPath);
        Console.WriteLine(pdfPath);
        Console.WriteLine(outputWithSummary);
    }
}