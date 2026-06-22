using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerHtmlDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Prepare a data source (DataTable) for the smart marker
                DataTable dt = new DataTable("Data");
                dt.Columns.Add("Name", typeof(string));
                dt.Rows.Add("Alice");
                dt.Rows.Add("Bob");
                dt.Rows.Add("Charlie");

                // Set the cell's HtmlString with a colored <span> that contains a smart marker
                // The smart marker &=Data.Name& will be replaced by each row's Name value during processing
                sheet.Cells["A1"].HtmlString = "<span style='color:#FF4500'>&=Data.Name&</span>";

                // Process the smart markers using WorkbookDesigner (correct API)
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dt);
                designer.Process();

                // Define output file path
                string outputPath = "SmartMarkerHtmlDemo.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}