// Title: Aspose.Cells .NET – Change shape text Latin (Western) font to Calibri without affecting FarEast font
// Description: Shows how to create a workbook, add a rectangle shape, and use the shape’s TextOptions to set the LatinName property to "Calibri" while leaving the FarEastName unchanged, then save the file.
// Keywords: Aspose.Cells | .NET | shape text font | LatinName Calibri | preserve FarEastName | TextOptions | rectangle shape | C# example
// Common Searches: Aspose.Cells set LatinName Calibri | keep FarEast font when changing shape text font | C# change western font in shape Aspose.Cells | modify TextOptions without altering Asian font | how to apply Calibri to shape text in Aspose.Cells .NET
// Developer Intent: Apply Calibri to the Western script of a shape’s text while retaining the existing Asian script font.
// Use Cases: Generate spreadsheets where English characters use Calibri but Japanese/Korean characters keep their default font | Automate template styling by updating only the Latin script font in diagram shapes | Prepare reports that require distinct fonts for Western and Far‑East text within the same shape
// AI Prompts: Write C# code using Aspose.Cells to set the LatinName of a shape’s TextOptions to "Calibri" and ensure FarEastName stays unchanged. | Explain steps to verify that the FarEast font was not modified after updating the Latin script font. | Show how to assign separate fonts for Latin and FarEast scripts in a single shape with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Shows how to create a workbook, add a rectangle shape, and use the shape’s TextOptions to set the LatinName property to "Calibri" while leaving the FarEastName unchanged, then save the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape that will contain text
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 100);
        shape.Text = "Western text example";

        // Obtain the TextOptions of the first paragraph inside the shape
        TextOptions textOptions = shape.TextBody.TextParagraphs[0].TextOptions;

        // Change the Latin (Western) font name to Calibri
        textOptions.LatinName = "Calibri";

        // Do NOT modify the FarEast font name; it remains whatever it was (default or previously set)
        Console.WriteLine("FarEastName unchanged: " + textOptions.FarEastName);

        // Save the workbook to a file
        workbook.Save("LatinFontCalibreDemo.xlsx");
    }
}
