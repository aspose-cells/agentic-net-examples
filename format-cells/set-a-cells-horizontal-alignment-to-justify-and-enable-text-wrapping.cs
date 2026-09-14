// Title: How to set a cell's horizontal alignment to Justify and enable text wrapping with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that applies justified horizontal alignment and turns on text wrapping for a specific cell using Aspose.Cells. | Demonstrate retrieving a cell's style, setting HorizontalAlignment to Justify, enabling text wrap, auto‑fitting the column, and saving the workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# set cell alignment to justify and wrap text | C# Aspose.Cells enable text wrapping for a particular cell | how to apply horizontal text justification in Excel using Aspose.Cells .NET | auto fit column after enabling wrap text with Aspose.Cells C# | apply cell style with justified alignment and wrap in Aspose.Cells for .NET
// Tags: horizontal text justification Aspose.Cells C# | cell text wrap Aspose.Cells .NET | auto‑fit column after wrap Aspose.Cells | modify cell style justification Aspose.Cells | apply cell alignment and wrap Aspose.Cells

using Aspose.Cells;

// Creates a workbook, inserts long text into cell A1, retrieves the cell's style, sets HorizontalAlignment to Justify, enables text wrapping, auto‑fits the column to show the wrapped content, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get the target cell (e.g., A1) and put some text
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("This is a long piece of text that should be justified and wrapped within the cell.");

        // Retrieve the cell's current style
        Style style = cell.GetStyle();

        // Set horizontal alignment to Justify
        style.HorizontalAlignment = TextAlignmentType.Justify;

        // Enable text wrapping
        style.IsTextWrapped = true;

        // Apply the modified style back to the cell
        cell.SetStyle(style);

        // Optionally auto‑fit the column to see the wrapping effect
        sheet.AutoFitColumn(0);

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
