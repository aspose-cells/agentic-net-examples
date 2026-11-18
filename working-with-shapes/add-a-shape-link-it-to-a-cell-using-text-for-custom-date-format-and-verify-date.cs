// Title: Create a rectangle shape linked to a cell and show a custom formatted date using the TEXT function in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# link a shape to a cell with a TEXT formula for custom date format | How to display a formatted date inside a rectangle shape using Aspose.Cells | Verify shape text after calling Workbook.CalculateFormula in Aspose.Cells | Add a shape and bind its text to cell A1 date value in Aspose.Cells .NET
// Tags: add rectangle shape with formula Aspose.Cells | shape text bound to cell via TEXT function | custom date format in shape Aspose.Cells | recalculate formulas for shape text C# | validate shape displayed text Aspose.Cells | save workbook containing linked shape .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example creates a workbook, writes a DateTime to cell A1, applies a short date style, adds a rectangle shape, sets the shape's text to a TEXT formula that formats the cell value as "dd-MMM-yyyy", recalculates formulas, compares the shape's displayed text with the expected formatted date, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Insert a date value into cell A1
            DateTime sampleDate = new DateTime(2023, 12, 25);
            worksheet.Cells["A1"].PutValue(sampleDate);

            // Apply a short date number format to the cell (optional)
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14; // built‑in short date format
            worksheet.Cells["A1"].SetStyle(dateStyle);

            // Add a rectangle shape to the worksheet
            // In newer Aspose.Cells versions AddShape returns the created Shape directly
            Shape shape = worksheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                5,   // upper left row
                0,   // upper left column
                5,   // top offset (pixels)
                0,   // left offset (pixels)
                200, // width (pixels)
                50   // height (pixels)
            );

            // Link the shape's text to cell A1 using TEXT for a custom date format
            shape.Text = "=TEXT(A1,\"dd-MMM-yyyy\")";

            // Recalculate formulas so the shape displays the evaluated text
            workbook.CalculateFormula();

            // Verify that the shape's displayed text matches the expected formatted date
            string expectedText = sampleDate.ToString("dd-MMM-yyyy");
            string actualText = shape.Text; // after calculation this holds the result

            Console.WriteLine($"Expected: {expectedText}");
            Console.WriteLine($"Actual  : {actualText}");
            Console.WriteLine($"Verification: {(expectedText == actualText ? "Passed" : "Failed")}");

            // Save the workbook (optional)
            string outputPath = "output.xlsx";
            try
            {
                // Ensure the directory exists before saving
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
