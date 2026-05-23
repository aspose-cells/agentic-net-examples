using System;
using System.Data;
using System.Globalization;
using System.IO;
using Aspose.Cells;

class SmartMarkerLocaleDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Set up headers
            ws.Cells["A1"].PutValue("Date");
            ws.Cells["B1"].PutValue("Amount");

            // Insert smart markers with format strings.
            // The date format will display month names according to the locale.
            // The number format will use the locale's list separator and decimal symbol.
            ws.Cells["A2"].PutValue("&=Date:dd MMMM yyyy");
            ws.Cells["B2"].PutValue("&=Amount:#,##0.00");

            // Prepare a simple data source for the smart markers
            DataTable dt = new DataTable();
            dt.Columns.Add("Date", typeof(DateTime));
            dt.Columns.Add("Amount", typeof(double));
            dt.Rows.Add(new DateTime(2023, 12, 31), 12345.67);

            // Create a WorkbookDesigner to process smart markers
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = wb
            };

            // Assign the data source (as a DataTable)
            designer.SetDataSource(dt);

            // Align the workbook's culture with the desired locale (French)
            wb.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Process the smart markers – the placeholders will be replaced using the French locale.
            designer.Process();

            // Define output path
            string outputPath = "SmartMarkerLocaleDemo.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the resulting workbook
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}