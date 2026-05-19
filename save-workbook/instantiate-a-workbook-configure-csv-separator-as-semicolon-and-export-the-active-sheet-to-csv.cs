using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    public class ExportActiveSheetToCsv
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (default format is Xlsx)
                Workbook workbook = new Workbook();

                // Access the first worksheet (active by default) and add sample data
                Worksheet activeSheet = workbook.Worksheets[0];
                activeSheet.Name = "ActiveSheet";
                activeSheet.Cells["A1"].PutValue("Name");
                activeSheet.Cells["B1"].PutValue("Age");
                activeSheet.Cells["A2"].PutValue("John");
                activeSheet.Cells["B2"].PutValue(30);
                activeSheet.Cells["A3"].PutValue("Alice");
                activeSheet.Cells["B3"].PutValue(25);

                // Add a second worksheet to demonstrate that only the active sheet will be exported
                Worksheet otherSheet = workbook.Worksheets.Add("OtherSheet");
                otherSheet.Cells["A1"].PutValue("Should not appear in CSV");

                // Ensure the first worksheet is the active one (optional, default is 0)
                workbook.Worksheets.ActiveSheetIndex = 0;

                // Configure CSV save options: separator set to semicolon
                TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    Separator = ';'
                };

                // Define output path
                string outputPath = "ActiveSheetExport.csv";

                // Save the active worksheet as CSV using the configured options
                workbook.Save(outputPath, csvOptions);

                Console.WriteLine($"Active worksheet exported to CSV with semicolon separator at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during export: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportActiveSheetToCsv.Run();
        }
    }
}