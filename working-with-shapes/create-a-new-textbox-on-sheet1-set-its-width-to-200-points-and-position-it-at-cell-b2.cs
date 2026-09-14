// Title: Create a TextBox shape on Sheet1 at cell B2 with a 200‑point width using Aspose.Cells for .NET
// AI Prompts: Write C# code that adds a TextBox shape to the first worksheet at cell B2, sets its width to 200 points, and saves the workbook with Aspose.Cells. | Generate a program that creates a new workbook, inserts a TextBox at B2 with a width of 200 points (height 100 points), assigns sample text, and writes the file to disk using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells how to place a TextBox at a specific cell in C# | set textbox width in points when adding shape with Aspose.Cells | C# Aspose.Cells add TextBox shape to Sheet1 at B2 | adjust dimensions of a TextBox shape using Aspose.Cells API | save workbook after inserting a TextBox with custom size Aspose.Cells
// Tags: add textbox shape Aspose.Cells C# | set textbox width points Aspose.Cells | position shape at cell B2 Aspose.Cells | textbox dimensions Excel Aspose.Cells | create shape on worksheet Aspose.Cells

using Aspose.Cells;

// The example creates a new workbook, adds a TextBox shape to Sheet1 at cell B2 with a width of 200 points (height 100 points), sets its text, and saves the file as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Get the first worksheet (Sheet1)
        var sheet = workbook.Worksheets[0];

        // Add a TextBox at cell B2 (row index 1, column index 1)
        // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
        var textBox = sheet.Shapes.AddTextBox(1, 1, 0, 0, 100, 200);

        // Set optional text inside the TextBox
        textBox.Text = "Sample TextBox";

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
