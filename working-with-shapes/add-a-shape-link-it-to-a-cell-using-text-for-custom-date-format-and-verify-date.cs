// Title: Add a rectangle shape linked to cell A1 via TEXT function for a custom date format and verify the shape text as a date using Aspose.Cells for .NET
// AI Prompts: Generate C# code that inserts a rectangle shape into a worksheet, sets its Text property to =TEXT(A1,"yyyy-MM-dd"), runs workbook.CalculateFormula(), and determines whether the shape's displayed text can be parsed into a valid DateTime. | Write a .NET example that writes today's date to cell A1, applies a standard date number format, links a rectangle shape's text to that cell using the TEXT function with a custom format, and validates the resulting string as a proper date.
// Common Searches: how to link a shape to a cell using TEXT formula in Aspose.Cells C# | validate shape text as date after linking to worksheet cell with Aspose.Cells | Aspose.Cells add rectangle shape and display formatted date from a cell | C# calculate formulas for shape text linked to a cell in Aspose.Cells
// Tags: Aspose.Cells add rectangle shape | shape text linked to cell using TEXT function | custom date format in shape text | validate displayed date from shape .NET | calculate formulas for linked shape Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates creating a workbook, inserting today's date into cell A1 with a standard number format, adding a rectangle shape, linking its Text property to the cell using the TEXT function for a yyyy-MM-dd format, calculating formulas, retrieving and validating the shape's text as a date, and saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a date value into cell A1
            DateTime today = DateTime.Today;
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(today);

            // Apply a standard date number format (Number format ID 14)
            Style style = cell.GetStyle();
            style.Number = 14;
            cell.SetStyle(style);

            // Add a rectangle shape to the worksheet
            // Parameters: type, upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle,
                5, 5, 0, 0, 50, 200);

            // Link the shape's text to cell A1 using TEXT with a custom date format
            shape.Text = "=TEXT(A1,\"yyyy-MM-dd\")";

            // Calculate all formulas so the shape displays the formatted date
            workbook.CalculateFormula();

            // Retrieve the displayed text from the shape
            string shapeText = shape.Text;

            // Verify that the shape text is a valid date in the expected format
            bool isValid = DateTime.TryParseExact(
                shapeText,
                "yyyy-MM-dd",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out DateTime parsedDate);

            Console.WriteLine($"Shape text: {shapeText}");
            Console.WriteLine($"Is valid date: {isValid}");
            if (isValid)
            {
                Console.WriteLine($"Parsed date: {parsedDate:d}");
            }

            // Save the workbook (optional)
            string outputPath = "ShapeLinkedDate.xlsx";

            // Ensure the directory exists before saving
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
