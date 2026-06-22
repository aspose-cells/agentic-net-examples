using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsStyleDemo
{
    public class ApplyStyleToDataSetRange
    {
        // Entry point for the console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // -------------------------------------------------
            // Example data – normally you would import a DataSet here.
            // For demonstration we fill a 5x3 block with sample values.
            // -------------------------------------------------
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Create a style using the Workbook.CreateStyle() method (rule)
            Style style = workbook.CreateStyle();

            // Set solid fill with light blue background
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = Color.LightBlue;

            // Make the font bold
            style.Font.IsBold = true;

            // Define the range that corresponds to the imported DataSet.
            // Here we use the same 5 rows x 3 columns we populated above.
            AsposeRange dataRange = cells.CreateRange(0, 0, 5, 3);

            // Apply the style to the entire range (rule: Range.SetStyle)
            dataRange.SetStyle(style);

            // Save the workbook (lifecycle rule: save)
            string outputPath = "StyledDataSetRange.xlsx";
            workbook.Save(outputPath);
        }
    }
}