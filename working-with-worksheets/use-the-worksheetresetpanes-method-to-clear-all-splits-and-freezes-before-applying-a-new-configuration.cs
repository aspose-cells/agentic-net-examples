using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class ResetPanesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Example: create a split and freeze panes to demonstrate existing configuration
                worksheet.Split();
                worksheet.FreezePanes("C3", 3, 3);

                // Reset panes: remove any split and unfreeze any frozen panes
                worksheet.RemoveSplit();      // Clears split window
                worksheet.UnFreezePanes();    // Unfreezes panes

                // Apply new pane configuration (e.g., freeze at cell B2)
                worksheet.FreezePanes("B2", 1, 1);

                // Save the workbook
                string outputPath = "ResetPanesDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ResetPanesDemo.Run();
        }
    }
}