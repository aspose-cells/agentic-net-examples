// Title: Insert a TextBox shape into a worksheet covering a specific cell range with Aspose.Cells for .NET (C#)
// Common Searches: how to insert a textbox shape between B2 and D5 using Aspose.Cells C# | Aspose.Cells C# add textbox to worksheet with pixel dimensions | set textbox font color and size in Aspose.Cells .NET | save workbook after adding shape with Aspose.Cells C# example | create output directory before saving workbook Aspose.Cells
// Tags: add textbox shape Aspose.Cells | textbox placement cell range Aspose.Cells | set textbox font properties Aspose.Cells | specify textbox pixel size Aspose.Cells | save workbook after inserting shape Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, accesses the first Worksheet, and adds a TextBox shape that spans cells B2 to D5 with a size of 150 × 80 pixels. It sets the textbox text, changes the font color to blue and size to 12, ensures the output directory exists, saves the workbook as Output.xlsx, and includes basic exception handling.
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

            // Define the cell coordinates for the textbox (e.g., from B2 to D5)
            int upperLeftRow = 1;      // Row index for B2 (zero‑based)
            int upperLeftColumn = 1;   // Column index for B2 (zero‑based)
            int lowerRightRow = 4;     // Row index for D5 (zero‑based)
            int lowerRightColumn = 3;  // Column index for D5 (zero‑based)

            // Width and height of the textbox in pixels (adjust as needed)
            int width = 150;
            int height = 80;

            // Add a textbox shape to the worksheet and obtain the TextBox object directly
            TextBox textbox = sheet.Shapes.AddTextBox(
                upperLeftRow,
                upperLeftColumn,
                lowerRightRow,
                lowerRightColumn,
                width,
                height);

            // Set the textbox content
            textbox.Text = "Hello, Aspose.Cells!";

            // Optional: format the textbox text
            textbox.Font.Color = Color.Blue;
            textbox.Font.Size = 12;

            // Prepare output path and ensure the directory exists
            string outputPath = "Output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
