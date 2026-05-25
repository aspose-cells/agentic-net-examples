using System;
using Aspose.Cells;

namespace AsposeCellsRichTextPortionDemo
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

                // Access a cell and assign a string value (rich text can be applied to any string)
                Cell cell = worksheet.Cells["A1"];
                cell.PutValue("Hello World! This is a test.");

                // Retrieve all rich‑text portions (FontSetting objects) from the cell
                FontSetting[] portions = cell.GetCharacters();

                // If the cell contains rich‑text portions, modify their font
                if (portions != null)
                {
                    foreach (FontSetting portion in portions)
                    {
                        // Set the desired font name, e.g., Arial
                        portion.Font.Name = "Arial";
                    }
                }

                // Save the workbook to a file
                workbook.Save("RichTextPortionModified.xlsx");
            }
            catch (Exception ex)
            {
                // Log or handle exceptions as needed
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}