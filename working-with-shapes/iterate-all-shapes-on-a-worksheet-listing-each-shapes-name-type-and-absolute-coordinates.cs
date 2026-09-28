// Title: Iterate through all shapes in an Excel worksheet and output each shape’s name, type, and absolute coordinates with Aspose.Cells for C#
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, loops over worksheet.Shapes, and prints each shape’s Name, Type, Left, Top, Width, and Height. | Create a reusable method that accepts a Worksheet object and returns a list of objects containing shape metadata (name, type, left, top, width, height) using Aspose.Cells. | Generate a script that extracts shape information from a workbook and writes the results to a CSV file, including Z‑order index along with coordinates.
// Common Searches: aspocells c# enumerate shapes on a worksheet and get their positions | how to read shape left and top values from an Excel file using Aspose.Cells .NET | retrieve shape type and name from a .xlsx workbook with Aspose.Cells in C# | list all drawing objects coordinates in Excel using Aspose.Cells API
// Tags: enumerate worksheet shapes Aspose.Cells C# | extract shape coordinates Excel Aspose.Cells | retrieve shape type and name Aspose.Cells | list shape bounds in .xlsx using C# | Aspose.Cells shape metadata extraction

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, accesses the first worksheet, iterates over every Shape object, and prints each shape’s name, type, and absolute position (left, top, width, height) to the console, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Retrieve shape properties
                string name = shape.Name;
                string type = shape.Type.ToString();
                double left = shape.Left;
                double top = shape.Top;
                double width = shape.Width;
                double height = shape.Height;

                // Output shape information
                Console.WriteLine($"Name: {name}, Type: {type}, Left: {left}, Top: {top}, Width: {width}, Height: {height}");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
