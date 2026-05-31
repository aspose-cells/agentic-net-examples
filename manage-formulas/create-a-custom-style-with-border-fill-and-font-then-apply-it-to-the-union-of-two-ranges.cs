using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class UnionRangeCustomStyleDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data in two separate ranges
                worksheet.Cells["A1"].PutValue("First");
                worksheet.Cells["B2"].PutValue(123);
                worksheet.Cells["D4"].PutValue("Second");
                worksheet.Cells["E5"].PutValue(456);

                // Create a UnionRange that covers A1:B2 and D4:E5
                UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A1:B2,D4:E5", 0);

                // Create a custom style with font, fill, and borders
                Style style = workbook.CreateStyle();

                // Font settings
                style.Font.Name = "Calibri";
                style.Font.Size = 12;
                style.Font.IsBold = true;
                style.Font.Color = Color.White;

                // Fill settings
                style.Pattern = BackgroundType.Solid;
                style.ForegroundColor = Color.DarkSlateBlue;

                // Border settings (apply same style to all four borders)
                style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
                style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
                style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
                style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;

                style.Borders[BorderType.TopBorder].Color = Color.Yellow;
                style.Borders[BorderType.BottomBorder].Color = Color.Yellow;
                style.Borders[BorderType.LeftBorder].Color = Color.Yellow;
                style.Borders[BorderType.RightBorder].Color = Color.Yellow;

                // Apply the custom style to the union range
                unionRange.SetStyle(style);

                // Define output file path
                string outputPath = "UnionRangeCustomStyleDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            UnionRangeCustomStyleDemo.Run();
        }
    }
}