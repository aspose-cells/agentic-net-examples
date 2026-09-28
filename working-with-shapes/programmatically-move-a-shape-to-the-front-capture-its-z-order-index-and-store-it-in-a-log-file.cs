// Title: Move a worksheet shape to the front and log its Z‑order index using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, brings a specific worksheet shape to the front by setting its ZOrderPosition, and appends the shape name and new Z‑order value to a text log. | Show how to programmatically reorder a shape on an Aspose.Cells worksheet, retrieve the resulting Z‑order index, write the information to a log file, and save the updated workbook.
// Common Searches: Aspose.Cells C# bring shape to front and get Z-order position | how to log shape Z-order after moving in an Excel file using Aspose.Cells | set highest ZOrderPosition for worksheet shape Aspose.Cells .NET example | append shape reorder details to a text file with Aspose.Cells | save workbook after changing shape order Aspose.Cells C#
// Tags: Aspose.Cells shape ZOrderPosition manipulation | move worksheet shape to front .NET | log shape Z-order index C# | append shape operation details to text file | save workbook after shape reorder Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads (or creates) an Excel workbook, accesses the first worksheet, moves the first shape to the front by assigning it the highest ZOrderPosition, records the shape's name and new Z‑order index to a text log, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string logPath = "shape_log.txt";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the collection of shapes on the worksheet
            ShapeCollection shapes = sheet.Shapes;

            // Ensure there is at least one shape to work with
            if (shapes.Count > 0)
            {
                // Use the first shape (or locate by name/id as needed)
                Shape shape = shapes[0];

                // Bring the shape to the front by setting its Z-order position to the highest index
                shape.ZOrderPosition = shapes.Count; // highest position

                // Capture the shape's Z-order index after moving it
                int zOrderIndex = shape.ZOrderPosition;

                // Prepare a log entry
                string logEntry = $"Shape '{shape.Name}' moved to front. Z-order index: {zOrderIndex}{Environment.NewLine}";

                // Append the log entry to a text file
                try
                {
                    File.AppendAllText(logPath, logEntry);
                }
                catch (Exception logEx)
                {
                    Console.WriteLine($"Logging error: {logEx.Message}");
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Save error: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
