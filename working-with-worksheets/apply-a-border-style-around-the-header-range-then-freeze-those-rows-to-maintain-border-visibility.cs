// Title: Add thin black borders to a header row and freeze the first row in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to apply a thin black border to a header range in Aspose.Cells C# | Aspose.Cells C# freeze first row after styling header | set border style on A1:D1 and freeze panes with Aspose.Cells .NET | C# code to create a header row with borders and freeze panes using Aspose.Cells | Aspose.Cells example for styling header and freezing rows
// Tags: cell border styling Aspose.Cells | freeze top worksheet row Aspose.Cells | header range formatting Excel .NET | apply style to range C# | worksheet freeze panes after styling

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsHeaderBorderAndFreeze
{
    // The example loads an existing workbook or creates a new one, defines the header range A1:D1, creates a thin black border style, applies the style to the header range, freezes the top row to keep the header visible while scrolling, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output files
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            Workbook workbook = null;

            try
            {
                // Load existing workbook if present; otherwise create a new one with sample data
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];

                    // Create a header row (A1:D1)
                    ws.Cells["A1"].PutValue("Header1");
                    ws.Cells["B1"].PutValue("Header2");
                    ws.Cells["C1"].PutValue("Header3");
                    ws.Cells["D1"].PutValue("Header4");

                    // Populate some sample data rows
                    for (int i = 2; i <= 10; i++)
                    {
                        ws.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                        ws.Cells[$"B{i}"].PutValue(i * 10);
                        ws.Cells[$"C{i}"].PutValue(i * 20);
                        ws.Cells[$"D{i}"].PutValue(i * 30);
                    }

                    Console.WriteLine("Created a new workbook with sample data.");
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Define the header range (first row, columns A through D)
                // Row and column indexes are zero‑based.
                int headerStartRow = 0;
                int headerStartColumn = 0;
                int headerRowCount = 1;
                int headerColumnCount = 4; // A, B, C, D

                AsposeRange headerRange = worksheet.Cells.CreateRange(
                    headerStartRow,
                    headerStartColumn,
                    headerRowCount,
                    headerColumnCount);

                // Create a style that contains a thin black border on all sides
                Style borderStyle = workbook.CreateStyle();
                borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

                borderStyle.Borders[BorderType.TopBorder].Color = Color.Black;
                borderStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
                borderStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
                borderStyle.Borders[BorderType.RightBorder].Color = Color.Black;

                // Apply the style to the header range
                StyleFlag applyAll = new StyleFlag { All = true };
                headerRange.ApplyStyle(borderStyle, applyAll);
                Console.WriteLine("Applied border style to the header range.");

                // Freeze the first row so the header remains visible while scrolling
                worksheet.FreezePanes(0, 0, 1, 0);
                Console.WriteLine("Frozen the first row to keep the header visible.");

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
