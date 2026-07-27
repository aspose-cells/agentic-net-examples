// Title: Set FarEast (CJK) and Latin Font Names for Shape Text in Aspose.Cells .NET
// Description: Shows how to add a rectangle shape to a worksheet, assign multilingual text, set the FarEast font to MS Mincho and the Latin font to Arial via Shape.TextOptions, and save the workbook.
// Keywords: Aspose.Cells | C# shape font | FarEastName | LatinName | MS Mincho | Arial | multilingual text | Excel shape | TextOptions | CJK font
// Common Searches: Aspose.Cells set CJK font for shape | How to use FarEastName in Aspose.Cells C# | Specify different fonts for Japanese and English text in Excel shape | Shape.TextOptions LatinName example | Change font of shape text Aspose.Cells .NET
// Developer Intent: Apply separate FarEast (CJK) and Latin fonts to the text of a shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Create a rectangle shape with Japanese characters displayed in MS Mincho and English characters in Arial. | Generate multilingual reports where shape labels need appropriate East Asian and Western fonts for readability. | Automate workbook creation that includes shapes containing mixed‑language text with font control per script.
// AI Prompts: Provide C# code to set FarEastName to MS Mincho and LatinName to Arial for a shape's text in Aspose.Cells. | Explain how Shape.TextOptions influences multilingual rendering in Excel shapes with Aspose.Cells .NET. | Show an example of applying different fonts to CJK and Latin characters inside a rectangle shape using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to add a rectangle shape to a worksheet, assign multilingual text, set the FarEast font to MS Mincho and the Latin font to Arial via Shape.TextOptions, and save the workbook.
class SetMultilingualFontDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 100);
        shape.Text = "Sample multilingual text";

        // Specify the FarEast (CJK) font name
        shape.TextOptions.FarEastName = "MS Mincho";

        // Specify the Latin (Western) font name
        shape.TextOptions.LatinName = "Arial";

        // Save the workbook to a file
        workbook.Save("MultilingualFontDemo.xlsx");
    }
}
