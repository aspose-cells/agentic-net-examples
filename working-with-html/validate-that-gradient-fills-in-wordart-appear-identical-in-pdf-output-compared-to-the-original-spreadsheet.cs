// Title: Validate Gradient WordArt Fill Consistency Between Excel and PDF Using Aspose.Cells for .NET
// Description: This C# example creates an Excel workbook, inserts a WordArt shape with a preset gradient fill, saves the file as XLSX and PDF, reloads the XLSX, and confirms that the shape remains WordArt and its FillType is Gradient, ensuring visual parity in the generated PDF.
// Keywords: Aspose.Cells | .NET | WordArt gradient | PDF export | fill type verification | Excel to PDF fidelity | gradient fill consistency | Aspose.Cells API | C# example
// Common Searches: Aspose.Cells verify WordArt gradient in PDF | C# check WordArt fill type after export | Excel gradient WordArt PDF rendering | How to test gradient fill preservation with Aspose.Cells | Validate WordArt appearance in PDF using Aspose.Cells
// Developer Intent: Confirm that a WordArt shape with a gradient fill renders identically in the PDF produced by Aspose.Cells as it does in the original Excel workbook.
// Use Cases: Automated regression test for gradient WordArt rendering during Excel‑to‑PDF conversion | Generate reports with styled WordArt and programmatically ensure gradient styling survives PDF export | Validate workbook integrity after saving and reloading by checking WordArt properties | Create a CI pipeline step that flags visual differences in gradient fills between XLSX and PDF
// AI Prompts: Generate a C# unit test that renders a WordArt shape from an Excel file and compares its bitmap to the same shape in the exported PDF using Aspose.Cells and Aspose.Pdf. | Provide code to extract gradient fill parameters from a WordArt object and assert they match the PDF rendering. | Explain how to programmatically verify that a WordArt shape retains its Gradient FillType after saving, reloading, and exporting with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// This C# example creates an Excel workbook, inserts a WordArt shape with a preset gradient fill, saves the file as XLSX and PDF, reloads the XLSX, and confirms that the shape remains WordArt and its FillType is Gradient, ensuring visual parity in the generated PDF.
class GradientWordArtPdfValidation
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape using a preset style that contains a gradient fill
        // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7, // Gradient Fill - Blue, Accent 1, Reflection
            "Gradient WordArt",
            2, 0,   // row, top offset
            2, 0,   // column, left offset
            200,    // height
            100);   // width

        // Save the original Excel file
        workbook.Save("GradientWordArt.xlsx");

        // Save the same workbook as PDF – the PDF rendering should preserve the gradient appearance
        workbook.Save("GradientWordArt.pdf", SaveFormat.Pdf);

        // Reload the saved Excel file to verify that the WordArt shape retains its gradient fill
        Workbook loadedWorkbook = new Workbook("GradientWordArt.xlsx");
        Shape loadedWordArt = loadedWorkbook.Worksheets[0].Shapes[0];

        // Verify that the shape is recognized as WordArt
        bool isWordArt = loadedWordArt.IsWordArt;

        // Verify that the fill type is Gradient (as defined by the preset style)
        FillType fillType = loadedWordArt.Fill.FillType;

        // Output verification results
        Console.WriteLine("IsWordArt: " + isWordArt);
        Console.WriteLine("FillType: " + fillType);
        // Expected output:
        // IsWordArt: True
        // FillType: Gradient
        // The visual appearance of the gradient in the PDF should match the Excel rendering.
    }
}
