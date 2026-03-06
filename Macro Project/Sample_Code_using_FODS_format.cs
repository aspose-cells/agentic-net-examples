using System;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsFodsDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sample";

            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);

            // Configure ODS save options for FODS format
            OdsSaveOptions saveOptions = new OdsSaveOptions();
            saveOptions.GeneratorType = OdsGeneratorType.LibreOffice;
            saveOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf12;

            // Save the workbook as a Flat XML ODS (FODS) file
            string fodsPath = "Sample.fods";
            workbook.Save(fodsPath, saveOptions);
            Console.WriteLine($"Workbook saved as FODS to {fodsPath}");

            // Load the FODS file with load options
            OdsLoadOptions loadOptions = new OdsLoadOptions();
            loadOptions.ApplyExcelDefaultStyleToHyperlink = true;

            Workbook loadedWorkbook = new Workbook(fodsPath, loadOptions);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            // Output the loaded data to verify correctness
            Console.WriteLine("Loaded data:");
            for (int row = 0; row <= 3; row++)
            {
                string name = loadedSheet.Cells[row, 0].StringValue;
                string score = loadedSheet.Cells[row, 1].StringValue;
                Console.WriteLine($"{name}\t{score}");
            }

            // Save the loaded workbook as a regular ODS file to demonstrate round‑trip
            loadedWorkbook.Save("RoundTrip.ods", SaveFormat.Ods);
        }
    }
}