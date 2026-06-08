using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ColumnFontSettingsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Access the first column (index 0)
                Column column = worksheet.Cells.Columns[0];

                // Create a new style object from the workbook
                Style columnStyle = workbook.CreateStyle();

                // Configure the font: size, color, and underline
                columnStyle.Font.Size = 14;                         // Font size
                columnStyle.Font.Color = Color.Red;                // Font color
                columnStyle.Font.Underline = FontUnderlineType.Single; // Underline

                // Apply the style to the entire column
                column.SetStyle(columnStyle);

                // Define output file path
                string outputPath = "ColumnFontSettings.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ColumnFontSettingsDemo.Run();
        }
    }
}