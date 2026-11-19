// Title: How to assign a custom Name to a TextBox shape in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set custom Name for worksheet TextBox shape | Assign unique identifier to Excel TextBox using Aspose.Cells API | Retrieve a TextBox by its Name property with Aspose.Cells in .NET | How to name a TextBox shape for later manipulation in Aspose.Cells | Aspose.Cells example: naming and accessing TextBox objects in a workbook
// Tags: Aspose.Cells TextBox Name property | C# set shape name Aspose.Cells | unique identifier for worksheet TextBox | retrieve TextBox by name Aspose.Cells | Excel shape naming with Aspose.Cells API | Aspose.Cells shape manipulation C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, adds a TextBox shape at a specified cell location, sets its displayed text, assigns a unique Name (e.g., "UniqueTextBox_001") for later identification, and saves the workbook to Output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the position of the textbox (rows and columns)
            int upperLeftRow = 2;      // 0‑based row index
            int upperLeftColumn = 1;   // 0‑based column index

            // Offsets in pixels from the upper‑left cell
            int upperLeftRowOffset = 0;
            int upperLeftColumnOffset = 0;

            // Size of the textbox in pixels
            int height = 100;
            int width = 200;

            // Add a textbox shape to the worksheet
            TextBox textBox = sheet.Shapes.AddTextBox(
                upperLeftRow, upperLeftColumn,
                upperLeftRowOffset, upperLeftColumnOffset,
                height, width);

            // Set the displayed text
            textBox.Text = "Hello, Aspose.Cells!";

            // Assign a unique name for later identification
            textBox.Name = "UniqueTextBox_001";

            // Define output file path
            string outputPath = "Output.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
