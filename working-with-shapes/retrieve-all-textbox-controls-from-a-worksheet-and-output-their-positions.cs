// Title: Aspose.Cells C# – List all TextBox shapes with their cell coordinates
// Description: Creates a workbook, adds sample TextBox shapes, iterates through the worksheet's TextBoxCollection, and prints each box's UpperLeftRow, UpperLeftColumn, LowerRightRow, and LowerRightColumn values. The file can be saved to verify the layout.
// Keywords: Aspose.Cells TextBox coordinates | C# enumerate TextBox shapes | retrieve TextBox cell positions | TextBoxCollection iteration | Aspose.Cells shape properties
// Common Searches: how to get textbox row and column indices in Aspose.Cells | list all text boxes in an Excel worksheet using .NET | Aspose.Cells get UpperLeftRow UpperLeftColumn of shapes | extract textbox coordinates from a workbook
// Developer Intent: Obtain every TextBox object on a worksheet and output its upper‑left and lower‑right row/column indices.
// Use Cases: Audit a spreadsheet to ensure text boxes are placed within a required area. | Generate a report that lists each text box’s exact location for documentation. | Re‑position text boxes programmatically based on their current coordinates.
// AI Prompts: Write C# code with Aspose.Cells that moves all text boxes to start at row 1, column 1 while keeping their original size. | Create a method that returns a collection of tuples containing UpperLeftRow, UpperLeftColumn, LowerRightRow, and LowerRightColumn for every TextBox in a worksheet. | Explain how to filter a TextBoxCollection to include only boxes that intersect a specified cell range.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample TextBox shapes, iterates through the worksheet's TextBoxCollection, and prints each box's UpperLeftRow, UpperLeftColumn, LowerRightRow, and LowerRightColumn values. The file can be saved to verify the layout.
class RetrieveTextBoxes
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Add sample text boxes to demonstrate retrieval
        int tbIndex1 = worksheet.TextBoxes.Add(2, 3, 100, 200);
        worksheet.TextBoxes[tbIndex1].Text = "Sample Box 1";

        int tbIndex2 = worksheet.TextBoxes.Add(10, 5, 150, 250);
        worksheet.TextBoxes[tbIndex2].Text = "Sample Box 2";

        // Retrieve all TextBox controls from the worksheet
        TextBoxCollection textBoxes = worksheet.TextBoxes;

        // Output the position of each TextBox
        for (int i = 0; i < textBoxes.Count; i++)
        {
            TextBox tb = textBoxes[i];
            Console.WriteLine($"TextBox {i}: UpperLeftRow={tb.UpperLeftRow}, UpperLeftColumn={tb.UpperLeftColumn}, LowerRightRow={tb.LowerRightRow}, LowerRightColumn={tb.LowerRightColumn}");
        }

        // Save the workbook (optional, to verify the added text boxes)
        workbook.Save("TextBoxesPositions.xlsx");
    }
}
