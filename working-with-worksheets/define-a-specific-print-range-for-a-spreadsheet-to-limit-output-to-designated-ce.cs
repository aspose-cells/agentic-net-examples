using System;
using Aspose.Cells;

class SetPrintAreaDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data in the worksheet
        sheet.Cells["A1"].PutValue("Header1");
        sheet.Cells["B1"].PutValue("Header2");
        sheet.Cells["C1"].PutValue("Header3");
        for (int i = 2; i <= 10; i++)
        {
            sheet.Cells[$"A{i}"].PutValue($"R{i - 1}C1");
            sheet.Cells[$"B{i}"].PutValue($"R{i - 1}C2");
            sheet.Cells[$"C{i}"].PutValue($"R{i - 1}C3");
        }

        // Define the print area (e.g., cells A1 to C5)
        sheet.PageSetup.PrintArea = "A1:C5";

        // Configure HTML save options to export only the defined print area
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportPrintAreaOnly = true,   // Export only the print area
            ExportGridLines = true        // Optional: include grid lines in the output
        };

        // Save the workbook; only the specified print area will be exported
        workbook.Save("PrintAreaOutput.html", saveOptions);
    }
}