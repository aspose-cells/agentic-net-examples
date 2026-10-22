// Title: How to left-align the first line and center-align the second line in an Aspose.Cells TextBox using C#
// AI Prompts: Generate C# code that adds a TextBox shape to a worksheet and uses HTML <div> tags to left‑align the first line and center‑align the second line. | Create an Excel workbook with Aspose.Cells where a TextBox contains two lines of text, each with a different alignment, and save the file.
// Common Searches: Aspose.Cells C# set left alignment for first line and center alignment for second line in a TextBox shape | using HTML div tags to format multiline text alignment in Aspose.Cells TextBox | how to apply different text alignments within the same TextBox in an Excel workbook with Aspose.Cells | C# Aspose.Cells example of mixed alignment inside a shape
// Tags: Aspose.Cells TextBox HTML alignment | C# Aspose.Cells multiline TextBox formatting | Excel shape mixed text alignment Aspose.Cells | Aspose.Cells set per-line alignment in TextBox | Aspose.Cells create TextBox with custom HTML

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // C# program that creates a workbook, adds a TextBox shape, sets its Text property with HTML <div> elements to left‑align the first line and center‑align the second line, and saves the result as AlignedTextBox.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define position and size for the TextBox
                int upperLeftRow = 2;
                int upperLeftColumn = 2;
                int upperLeftRowOffset = 0;   // pixel offset from the upper‑left row
                int upperLeftColumnOffset = 0; // pixel offset from the upper‑left column
                int height = 100; // height in points
                int width = 200;  // width in points

                // Add a TextBox shape to the worksheet
                TextBox textBox = sheet.Shapes.AddTextBox(
                    upperLeftRow, upperLeftColumn,
                    upperLeftRowOffset, upperLeftColumnOffset,
                    height, width);

                // Set HTML content with alignment
                textBox.Text = "<div style='text-align:left'>First line</div>" +
                               "<div style='text-align:center'>Second line</div>";

                // Save the workbook
                string outputPath = "AlignedTextBox.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
