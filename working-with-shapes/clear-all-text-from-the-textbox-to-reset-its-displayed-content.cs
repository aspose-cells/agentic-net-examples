// Title: How to clear text from a TextBox shape in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove text from TextBox shape in worksheet | C# clear all textbox contents in an Excel file with Aspose.Cells | Programmatically reset textbox text in an Aspose.Cells workbook | Iterate worksheet shapes and clear TextBox text using Aspose.Cells .NET API | How to delete displayed text of a TextBox shape in Excel via Aspose.Cells
// Tags: Aspose.Cells clear textbox text C# | Aspose.Cells iterate worksheet shapes | MsoDrawingType TextBox Aspose.Cells | reset shape text Excel .NET | Aspose.Cells remove textbox content | C# clear textbox shape Excel workbook

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsClearTextbox
{
    // The example creates a workbook, adds a TextBox shape with sample text, saves the file, then iterates over all shapes on the worksheet, identifies those whose MsoDrawingType is TextBox, sets their Text property to an empty string to clear the displayed content, reports the number of cleared shapes, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for the demonstration files
            const string beforeClearPath = "BeforeClear.xlsx";
            const string afterClearPath = "AfterClear.xlsx";

            try
            {
                // ------------------------------------------------------------
                // 1. Create a new workbook and add a textbox with some text.
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a textbox shape to the worksheet.
                // Parameters: type, upperLeftRow, upperLeftColumn, top, left, width, height
                Shape textBox = sheet.Shapes.AddShape(
                    MsoDrawingType.TextBox, // Shape type
                    2,                      // Upper left row
                    2,                      // Upper left column
                    0,                      // Top offset (in pixels)
                    0,                      // Left offset (in pixels)
                    200,                    // Width (in pixels)
                    50);                    // Height (in pixels)

                // Set initial text for the textbox.
                textBox.Text = "Sample text inside the textbox.";
                Console.WriteLine("Created textbox with text: \"{0}\"", textBox.Text);

                // Save the workbook before clearing the text.
                workbook.Save(beforeClearPath, SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved as \"{0}\" (contains original textbox text).", beforeClearPath);

                // ------------------------------------------------------------
                // 2. Clear all text from every textbox shape in the worksheet.
                // ------------------------------------------------------------
                int clearedCount = 0;
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Identify textbox shapes via MsoDrawingType.
                        if (shape.MsoDrawingType == MsoDrawingType.TextBox)
                        {
                            // Clear the text. Setting to empty string removes displayed content.
                            shape.Text = string.Empty;
                            clearedCount++;
                        }
                    }
                    catch (Exception exShape)
                    {
                        Console.WriteLine($"Error processing shape '{shape.Name}': {exShape.Message}");
                    }
                }

                Console.WriteLine("Cleared text from {0} textbox shape(s).", clearedCount);

                // Save the workbook after clearing the text.
                workbook.Save(afterClearPath, SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved as \"{0}\" (textbox text cleared).", afterClearPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
