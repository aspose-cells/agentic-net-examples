// Title: Position a TextBox at exact point coordinates in a PDF generated with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add textbox at exact X Y points before saving as PDF | set textbox top and left offsets in points using Aspose.Cells AddTextBox | C# place custom text at (150,250) points in worksheet and export to PDF with Aspose.Cells | how to remove textbox border and format font when positioning shape in Aspose.Cells PDF output | define precise shape size and location in points for PDF generation with Aspose.Cells
// Tags: addtextbox exact point coordinates Aspose.Cells | position shape in points Aspose.Cells | export worksheet to PDF with positioned textbox | custom text placement PDF Aspose.Cells | remove textbox border Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new Workbook, adds a TextBox shape at X=150 and Y=250 points with defined width and height, sets custom text and font formatting, removes the border, and saves the worksheet as a PDF.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define position and size in points (1 point = 1/72 inch)
            int left = 150;   // X offset (points)
            int top = 250;    // Y offset (points)
            int width = 300;  // Width (points)
            int height = 100; // Height (points)

            // Add a TextBox shape to the worksheet.
            // Parameters: upperLeftRow, upperLeftColumn, top, left, height, width
            TextBox textBox = sheet.Shapes.AddTextBox(0, 0, top, left, height, width);

            // Set the custom text
            textBox.Text = "Precise custom text placed at (150,250) points";

            // Optional: format the text (font, size, color)
            textBox.Font.Name = "Arial";
            textBox.Font.Size = 12;
            textBox.Font.Color = Color.Black;

            // Optional: remove the border for a cleaner look
            textBox.Line.Weight = 0;

            // Save the workbook as PDF (lifecycle rule: save)
            workbook.Save("CustomText.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
