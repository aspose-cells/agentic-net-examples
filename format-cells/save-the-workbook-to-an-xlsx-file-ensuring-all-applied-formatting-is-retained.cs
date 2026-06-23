using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class SaveWorkbookWithFormattingDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                using (Workbook workbook = new Workbook())
                {
                    Worksheet worksheet = workbook.Worksheets[0];

                    // Add sample data
                    worksheet.Cells["A1"].PutValue("Product");
                    worksheet.Cells["B1"].PutValue("Price");
                    worksheet.Cells["A2"].PutValue("Apple");
                    worksheet.Cells["B2"].PutValue(1.25);
                    worksheet.Cells["A3"].PutValue("Banana");
                    worksheet.Cells["B3"].PutValue(0.75);

                    // Define header style
                    Style headerStyle = workbook.CreateStyle();
                    headerStyle.Font.IsBold = true;
                    headerStyle.ForegroundColor = Color.LightGray;
                    headerStyle.Pattern = BackgroundType.Solid;

                    // Apply style to header range
                    AsposeRange headerRange = worksheet.Cells.CreateRange("A1:B1");
                    headerRange.ApplyStyle(headerStyle, new StyleFlag { All = true });

                    // Auto‑fit columns for proper layout
                    worksheet.AutoFitColumns();

                    // Save the workbook
                    string outputPath = "FormattedOutput.xlsx";
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved as '{outputPath}' with formatting retained.");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Program entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            SaveWorkbookWithFormattingDemo.Run();
        }
    }
}