using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsTwoDimArrayDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Two‑dimensional array of strings (including header row)
                object[,] data = new object[,]
                {
                    { "Product", "Price", "Quantity" },   // Header row
                    { "Apple",   "1.20",  "50" },
                    { "Banana",  "0.80",  "100" },
                    { "Orange",  "1.00",  "75" }
                };

                // Import the array into the worksheet starting at cell A1
                // The last parameter prevents automatic conversion of strings to numbers
                cells.ImportTwoDimensionArray(data, 0, 0, false);

                // Create a custom style for the header row
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Font.IsBold = true;
                headerStyle.Font.Color = Color.Blue;
                headerStyle.Font.Size = 12;

                // Specify which style attributes to apply (font related only)
                StyleFlag headerFlag = new StyleFlag
                {
                    FontBold = true,
                    FontColor = true,
                    FontSize = true,
                    FontName = true
                };

                // Apply the style to the first row (header row)
                AsposeRange headerRange = cells.CreateRange(0, 0, 1, data.GetLength(1));
                headerRange.ApplyStyle(headerStyle, headerFlag);

                // Save the workbook
                string outputPath = "TwoDimArrayWithHeaderStyle.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}