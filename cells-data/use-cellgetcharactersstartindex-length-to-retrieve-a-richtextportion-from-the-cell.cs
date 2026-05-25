using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RetrieveRichTextPortionDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Put a string value into cell A1
                cells["A1"].PutValue("HelloWorld");

                // Retrieve a rich text portion (FontSetting) from the cell.
                // Characters(startIndex, length) returns a FontSetting object that
                // represents the specified range of characters.
                FontSetting richPortion = cells["A1"].Characters(5, 5); // "World"

                // Modify the retrieved portion's formatting
                richPortion.Font.IsBold = true;
                richPortion.Font.Color = Color.Blue;
                richPortion.Font.Size = 14;

                // Define output file path
                string outputPath = "RetrieveRichTextPortionDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            RetrieveRichTextPortionDemo.Run();
        }
    }
}