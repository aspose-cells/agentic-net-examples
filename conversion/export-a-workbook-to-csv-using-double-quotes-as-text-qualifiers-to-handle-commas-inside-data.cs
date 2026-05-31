using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate cells with data that includes commas
            cells["A1"].PutValue("Name");
            cells["B1"].PutValue("Description");
            cells["A2"].PutValue("Widget");
            cells["B2"].PutValue("Small, red, plastic");   // contains commas
            cells["A3"].PutValue("Gadget");
            cells["B3"].PutValue("Large, blue, metal");    // contains commas

            // Configure CSV save options
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Separator = ',',                     // Use comma as delimiter
                Encoding = Encoding.UTF8,            // UTF-8 encoding
                QuoteType = TxtValueQuoteType.Always // Always enclose fields in double quotes
            };

            // Save the workbook as CSV with the specified options
            string outputPath = "ExportedData.csv";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook exported to CSV with double‑quote qualifiers: {outputPath}");
        }
    }
}