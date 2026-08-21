// Title: Add a Rectangle Shape with an External Hyperlink in Excel using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a new Workbook, insert a rectangle shape at row 2/column 2, attach an external URL (https://www.example.com/) to the shape with AddHyperlink, and save the file as ShapeWithExternalHyperlink.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# shape hyperlink | Excel rectangle shape link | Add clickable shape Aspose.Cells | external URL shape .NET | C# Aspose.Cells tutorial | Excel hyperlink shape example
// Common Searches: Aspose.Cells add hyperlink to shape C# | Create clickable rectangle in Excel with Aspose | How to link a shape to a website using Aspose.Cells | C# code for shape hyperlink in Excel workbook
// Developer Intent: Insert a shape into a worksheet and bind it to an external web address.
// Use Cases: Add a company logo that opens the corporate site when clicked. | Create a call‑to‑action button linking to a marketing landing page. | Provide a quick‑access icon that launches an online help portal.
// AI Prompts: Write C# code with Aspose.Cells to add a circular shape that opens https://www.example.com/ when clicked. | Show how to set a shape’s hyperlink to open in a new browser tab using Aspose.Cells for .NET. | Give an example of adding multiple shapes, each with a different external URL, in the same worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsHyperlinkExample
{
    // Demonstrates how to create a new Workbook, insert a rectangle shape at row 2/column 2, attach an external URL (https://www.example.com/) to the shape with AddHyperlink, and save the file as ShapeWithExternalHyperlink.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet (row, column, width, height, upper left row, upper left column)
            // Here we place the shape at row 2, column 2 with a size of 100x100 pixels
            Shape shape = worksheet.Shapes.AddRectangle(2, 2, 100, 100, 0, 0);

            // Attach a hyperlink to the shape that points to an external website
            shape.AddHyperlink("https://www.example.com/");

            // Save the workbook to a file
            workbook.Save("ShapeWithExternalHyperlink.xlsx");
        }
    }
}
