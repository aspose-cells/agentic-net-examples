using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetColumnWidthPixelAfterAutoFitDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data in columns A to E (indexes 0‑4)
                cells["A1"].PutValue("Short");
                cells["B1"].PutValue("A bit longer text");
                cells["C1"].PutValue("This is a considerably longer piece of text");
                cells["D1"].PutValue("Medium length");
                cells["E1"].PutValue("Tiny");

                // Auto‑fit the columns to let Aspose.Cells determine the optimal width
                sheet.AutoFitColumns(0, 4);

                // Desired exact width in pixels for each column
                int desiredPixelWidth = 120;

                // Apply the exact pixel width to each column after auto‑fit
                for (int col = 0; col <= 4; col++)
                {
                    // Retrieve the width that AutoFitColumns assigned (in pixels)
                    int currentPixelWidth = cells.GetColumnWidthPixel(col);
                    Console.WriteLine($"Column {col} auto‑fit width: {currentPixelWidth} pixels");

                    // Set the column width to the desired pixel value
                    cells.SetColumnWidthPixel(col, desiredPixelWidth);
                }

                // Define output file path
                string outputPath = "SetColumnWidthPixelAfterAutoFitDemo.xlsx";

                // Save the workbook (overwrite if it already exists)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetColumnWidthPixelAfterAutoFitDemo.Run();
        }
    }
}