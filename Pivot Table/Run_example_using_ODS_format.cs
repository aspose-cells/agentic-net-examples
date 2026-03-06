using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsOdsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1500);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(2000);

            // Create ODS save options using the default constructor
            OdsSaveOptions saveOptions = new OdsSaveOptions();
            // Set the generator type (LibreOffice)
            saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;
            // Set the ODF strict version to 1.2
            saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf12;

            // Save the workbook as ODS with the specified options
            string odsPath = "SampleOutput.ods";
            workbook.Save(odsPath, saveOptions);
            Console.WriteLine($"Workbook saved to {odsPath}");

            // Load the saved ODS file with load options
            OdsLoadOptions loadOptions = new OdsLoadOptions();
            // Apply Excel default style to hyperlinks (example property)
            loadOptions.ApplyExcelDefaultStyleToHyperlink = true;

            Workbook loadedWorkbook = new Workbook(odsPath, loadOptions);
            // Verify a cell value after loading
            string product = loadedWorkbook.Worksheets[0].Cells["A2"].StringValue;
            Console.WriteLine($"Loaded product from cell A2: {product}");
        }
    }
}