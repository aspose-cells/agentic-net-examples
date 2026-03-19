using System;
using Aspose.Cells;

namespace AsposeCellsExportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add header values
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");

            // Apply bold style to the header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            StyleFlag flag = new StyleFlag();
            flag.FontBold = true;
            sheet.Cells.CreateRange("A1:C1").ApplyStyle(headerStyle, flag);

            // Add sample data rows
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(0.5);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["C3"].PutValue(0.3);

            // Configure save options (disable cell name export for better performance)
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.ExportCellName = false;

            // Save the workbook as an XLSX file
            string outputPath = "ExportedData.xlsx";
            workbook.Save(outputPath, saveOptions);

            // Output result to console
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}