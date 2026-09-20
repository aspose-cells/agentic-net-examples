// Title: How to Get TopMarginPt, LeftMarginPt, BottomMarginPt, and RightMarginPt of a Shape (TextBox) in an Excel Workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that loads an Excel file, locates the first shape, checks if it is a TextBox, and attempts to read its TopMarginPt, LeftMarginPt, BottomMarginPt, and RightMarginPt values, handling the case where the API does not expose these properties. | Create a C# example that safely casts a worksheet shape to a TextBox, retrieves any available text‑margin information, and logs a clear message when margin properties are unavailable in the current Aspose.Cells version. | Write a robust C# routine that enumerates all shapes in a worksheet, identifies TextBox shapes, and reports their text‑margin settings or indicates that the margin API is missing, including proper file‑existence checks and exception handling.
// Common Searches: Aspose.Cells C# read text box margin points in Excel | Get TopMarginPt of a shape using Aspose.Cells .NET | How to check if a shape supports margins in Aspose.Cells | C# Aspose.Cells shape margin properties not available | Retrieve left and right margin values from Excel TextBox with Aspose.Cells
// Tags: Aspose.Cells shape margin extraction limitation | C# retrieve TextBox margin settings in Excel | unsupported shape margin API Aspose.Cells | Excel worksheet shape margin access .NET | detect missing margin properties Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel workbook, verifies the file, accesses the first worksheet, checks for shapes, attempts to cast the first shape to a TextBox, and demonstrates that Aspose.Cells does not expose TopMarginPt, LeftMarginPt, BottomMarginPt, or RightMarginPt properties, providing clear messages and robust error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the worksheet
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the first worksheet.");
                return;
            }

            // Get the first shape on the worksheet
            Shape shape = sheet.Shapes[0];

            // Some shapes (e.g., TextBox) expose text‑related properties.
            // Cast to TextBox if possible.
            if (shape is TextBox textBox)
            {
                // Aspose.Cells TextBox does not expose margin properties directly.
                // If needed, you can work with other text formatting options here.
                Console.WriteLine("The selected shape is a TextBox, but margin properties are not available in this API version.");
            }
            else
            {
                Console.WriteLine("The selected shape does not support text margins (not a TextBox).");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
