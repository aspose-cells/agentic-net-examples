using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ApplyUniformColumnWidth
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and its cells collection
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Apply a uniform column width (in character units) to all columns
                cells.StandardWidth = 20.0; // Example width; adjust as needed

                // Verify the setting by outputting the standard width and the actual width of the first column
                Console.WriteLine("Standard Width set to: " + cells.StandardWidth);
                Console.WriteLine("Column 0 actual width: " + cells.GetColumnWidth(0));

                // Save the workbook to a file
                string outputPath = "UniformColumnWidth.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyUniformColumnWidth.Run();
        }
    }
}