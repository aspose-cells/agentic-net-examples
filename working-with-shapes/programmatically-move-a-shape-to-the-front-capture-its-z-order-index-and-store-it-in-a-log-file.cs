// Title: Move an Excel worksheet shape to the front, capture its Z‑order index, and log the change using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that selects a shape on a worksheet, moves it to the topmost Z‑order, and returns the new ZOrderPosition. | Create a method that updates a shape's ZOrderPosition to the highest value, records the shape name and new index, and appends the information to a text log. | Show how to iterate through all shapes in a worksheet to find the maximum Z‑order, then set a specific shape's ZOrderPosition to max+1 using Aspose.Cells.
// Common Searches: how to bring a shape to front in Excel using Aspose.Cells C# | retrieve and log Z-order of worksheet drawing objects with Aspose.Cells .NET | Aspose.Cells move shape to top layer and write details to a file | C# Aspose.Cells example for changing shape ZOrderPosition | log shape position changes in Excel workbook using Aspose.Cells
// Tags: Aspose.Cells move shape to front | C# Aspose.Cells ZOrderPosition | log shape changes to text file | Excel worksheet shape Z-order manipulation | Aspose.Cells shape ordering example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads an Excel workbook, moves the first shape on the first worksheet to the front by setting its ZOrderPosition to the highest existing value plus one, captures the shape name and new Z‑order index, appends this information to a log file, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string logPath = "shape_log.txt";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one shape on the sheet
            if (sheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found on the worksheet.");
                return;
            }

            // Get the first shape (or select by index/name as needed)
            Shape shape = sheet.Shapes[0];

            // Determine the highest Z-order position among all shapes
            int maxZOrder = 0;
            foreach (Shape s in sheet.Shapes)
            {
                if (s.ZOrderPosition > maxZOrder)
                    maxZOrder = s.ZOrderPosition;
            }

            // Move the shape to the front by setting its Z-order to max+1
            shape.ZOrderPosition = maxZOrder + 1;

            // Capture the current Z-order index after moving to front
            int zOrderIndex = shape.ZOrderPosition;

            // Prepare log entry
            string logEntry = $"Shape '{shape.Name}' moved to front. Z-order index: {zOrderIndex}{Environment.NewLine}";

            // Append the log entry to the log file
            File.AppendAllText(logPath, logEntry);

            // Save the modified workbook
            workbook.Save(outputPath);

            Console.WriteLine("Shape moved to front, Z-order captured, and log written successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
