// Title: Insert a TextBox shape with preset text into the first worksheet using Aspose.Cells for .NET
// AI Prompts: Create a new workbook, add a TextBox shape at row 2, column 2 with a width of 200 px and height of 100 px, set its caption to "Hello, Aspose.Cells!", and save the file as Output.xlsx. | Programmatically place a TextBox on the first worksheet, define its pixel dimensions, assign initial text, and export the workbook in XLSX format with Aspose.Cells.
// Common Searches: how to programmatically add a textbox to an Excel sheet with Aspose.Cells .NET | setting default text for a textbox shape in an Aspose.Cells workbook | specify textbox position using row and column indices in Aspose.Cells | saving an Excel file that contains a textbox shape using Aspose.Cells
// Tags: insert textbox shape Aspose.Cells | set textbox caption .NET | textbox position row column Aspose.Cells | textbox pixel dimensions Aspose.Cells | save workbook xlsx with shapes Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts a TextBox shape at row 2/column 2 with a 200 × 100 pixel size, assigns the text "Hello, Aspose.Cells!", and saves the result as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define position and size for the TextBox
        int upperLeftRow = 2;      // Row index (0‑based)
        int upperLeftColumn = 2;   // Column index (0‑based)
        int top = 0;               // Pixels from the top of the cell
        int left = 0;              // Pixels from the left of the cell
        int height = 100;          // Height in pixels
        int width = 200;           // Width in pixels

        // Add a TextBox shape to the worksheet
        Shape textBox = sheet.Shapes.AddShape(
            MsoDrawingType.TextBox,
            upperLeftRow,
            upperLeftColumn,
            top,
            left,
            height,
            width);

        // Set the initial text of the TextBox
        textBox.Text = "Hello, Aspose.Cells!";

        // Save the workbook to a file
        workbook.Save("Output.xlsx", SaveFormat.Xlsx);
    }
}
