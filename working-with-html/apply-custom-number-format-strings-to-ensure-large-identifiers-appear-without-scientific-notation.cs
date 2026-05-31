using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDemo
{
    class LargeIdentifierNumberFormatDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate cells with large numeric identifiers
                cells["A1"].PutValue("ID");
                cells["A2"].PutValue(1234567890123L);
                cells["A3"].PutValue(9876543210987L);
                cells["A4"].PutValue(5555555555555L);

                // Create a custom number format style that forces full integer display (no scientific notation)
                Style style = workbook.CreateStyle();
                style.Custom = "0"; // plain integer format

                // Apply only the number format to the range A2:A4 using a StyleFlag
                StyleFlag flag = new StyleFlag();
                flag.NumberFormat = true;
                Aspose.Cells.Range range = cells.CreateRange(1, 0, 3, 1); // rows 2‑4, column A
                range.ApplyStyle(style, flag);

                // Add a column chart to visualize the identifiers
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("A2:A4", true);               // values
                chart.NSeries.CategoryData = "A2:A4";           // categories (same as values for illustration)

                // Enable data labels and set the same custom number format for them
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;
                series.DataLabels.NumberFormat = "0";

                // Ensure the value axis tick labels also use the custom format
                chart.ValueAxis.TickLabels.NumberFormat = "0";

                // Save the workbook
                string outputPath = "LargeIdentifierNumberFormatDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            LargeIdentifierNumberFormatDemo.Run();
        }
    }
}