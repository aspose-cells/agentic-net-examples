using System;
using System.Data;
using System.IO;
using Aspose.Cells;   // Aspose.Cells namespace

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Insert a smart marker into cell A1 and wrap it with HTML <b> tags.
            // The HtmlString property preserves the HTML formatting after processing.
            worksheet.Cells["A1"].HtmlString = "<b>&lt;#Name#&gt;</b>";

            // Prepare a simple data source for the smart marker.
            DataTable table = new DataTable();
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("John Doe");

            // Process the smart marker using WorkbookDesigner (alternative to SmartMarkerProcessor).
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(table);
            designer.Process();   // processes all worksheets in the workbook

            // Define output path and ensure the directory exists
            string outputPath = "SmartMarkerBold.html";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to an HTML file to verify the bold rendering.
            workbook.Save(outputPath, SaveFormat.Html);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}