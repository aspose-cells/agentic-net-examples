using System;
using Aspose.Cells;

namespace AsposeCellsStyleExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a new style using the workbook's CreateStyle method
                Style fontStyle = workbook.CreateStyle();

                // Set the desired font family and size
                fontStyle.Font.Name = "Calibri";
                fontStyle.Font.Size = 14;

                // Define a StyleFlag to indicate which style attributes should be applied
                StyleFlag flag = new StyleFlag
                {
                    FontName = true,   // Apply font name
                    FontSize = true    // Apply font size
                };

                // Apply the style to a specific range (A1:C3) using the flag
                // Resolve ambiguity with System.Range by using the fully qualified name
                Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1", "C3");
                range.ApplyStyle(fontStyle, flag);

                // Optionally put some sample values in the range to see the style effect
                worksheet.Cells["A1"].PutValue("Cell A1");
                worksheet.Cells["B2"].PutValue("Cell B2");
                worksheet.Cells["C3"].PutValue("Cell C3");

                // Save the workbook to a file
                string outputPath = "StyledWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}