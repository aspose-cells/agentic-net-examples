using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsHeaderBoldDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate header cells (first row) with sample text
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Price");
                cells["C1"].PutValue("Quantity");

                // Create a style with bold font
                Style boldStyle = workbook.CreateStyle();
                boldStyle.Font.IsBold = true;

                // Define a StyleFlag to apply only the font bold attribute
                StyleFlag boldFlag = new StyleFlag
                {
                    FontBold = true
                };

                // Apply the bold style to the header range (first row, columns A to C)
                AsposeRange headerRange = cells.CreateRange(0, 0, 1, 3); // row 0, column 0, 1 row, 3 columns
                headerRange.ApplyStyle(boldStyle, boldFlag);

                // Save the workbook
                string outputPath = "HeaderBold.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Header cells have been styled with bold font and saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}