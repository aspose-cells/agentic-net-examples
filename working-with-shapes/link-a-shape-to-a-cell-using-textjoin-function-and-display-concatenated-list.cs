// Title: Create a textbox shape linked to a cell range using a TEXTJOIN formula in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add textbox shape with TEXTJOIN formula | How to bind a shape to a cell range and concatenate values using Aspose.Cells | Set TEXTJOIN function on a shape's text in a .NET workbook | Display concatenated list from A1:A5 in an Excel textbox using Aspose.Cells | Assign formula to textbox shape in Aspose.Cells .NET example
// Tags: shape text formula using TEXTJOIN | link shape to cell range Aspose.Cells | insert shape into worksheet .NET | concatenate range values in Excel shape | format shape font Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, populates cells A1 through A5 with fruit names, adds a textbox shape at row 5 column 0, sets its text to a TEXTJOIN formula that concatenates the range A1:A5, applies bold 12‑point font styling, and saves the file as LinkedShape_TextJoin.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Fill sample data in column A (A1:A5)
                sheet.Cells["A1"].PutValue("Apple");
                sheet.Cells["A2"].PutValue("Banana");
                sheet.Cells["A3"].PutValue("Cherry");
                sheet.Cells["A4"].PutValue("Date");
                sheet.Cells["A5"].PutValue("Elderberry");

                // Add a textbox shape.
                // Parameters: upperLeftRow, upperLeftColumn, top, left, height, width
                Shape shape = sheet.Shapes.AddTextBox(5, 0, 0, 0, 50, 200);

                // Set the shape's text to a TEXTJOIN formula that concatenates A1:A5
                shape.Text = "=TEXTJOIN(\", \", TRUE, A1:A5)";

                // Format the shape (font size, bold)
                shape.Font.Size = 12;
                shape.Font.IsBold = true;
                // Optional: horizontal alignment (uncomment if the enum is available)
                // shape.TextHorizontalAlignment = TextAlignmentType.Center;

                // Define output file path
                string outputPath = "LinkedShape_TextJoin.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
