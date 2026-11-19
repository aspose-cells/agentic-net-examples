// Title: How to add a clickable hyperlink to a textbox shape in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add hyperlink to textbox shape in Excel worksheet | set URL for textbox shape click event using Aspose.Cells .NET | create Excel textbox that opens a web page with Aspose.Cells API | assign Hyperlink.Address to shape in Aspose.Cells C# example | add clickable text box to .xlsx using Aspose.Cells for .NET
// Tags: Aspose.Cells textbox shape hyperlink | C# set shape hyperlink address | Excel worksheet add clickable textbox | Aspose.Cells save workbook with hyperlink | Hyperlink.Address property Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new workbook, inserts a textbox shape at a specified cell location, sets its text, assigns a Hyperlink.Address pointing to a web URL, and saves the workbook as an .xlsx file using Aspose.Cells for .NET.
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

            // Define textbox position and size (row/column indices are zero‑based)
            int upperLeftRow = 2;
            int upperLeftColumn = 2;
            int top = 5;    // Pixels from the top of the cell
            int left = 5;   // Pixels from the left of the cell
            int height = 100;
            int width = 200;

            // Add a textbox shape to the worksheet
            Shape textbox = sheet.Shapes.AddTextBox(upperLeftRow, upperLeftColumn, top, left, height, width);
            textbox.Text = "Click here to visit Aspose";

            // Set hyperlink for the textbox shape
            Hyperlink shapeLink = textbox.Hyperlink; // read‑only property returns a Hyperlink object
            shapeLink.Address = "https://www.aspose.com";

            // Save the workbook
            string outputFile = "HyperlinkTextbox.xlsx";
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputFile)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
