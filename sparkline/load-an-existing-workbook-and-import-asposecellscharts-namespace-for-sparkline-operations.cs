// Title: C# example: Load an Excel workbook and add a line sparkline with custom high/low colors using Aspose.Cells
// Common Searches: how to create a line sparkline in an existing Excel file using Aspose.Cells for .NET | Aspose.Cells C# load workbook and add sparkline group with custom colors | sample code to add a sparkline to cell G1 from range A1:E1 with Aspose.Cells | set high‑point green and low‑point red markers for a sparkline in C# using Aspose.Cells | add a sparkline when source data may be missing in an Excel worksheet with Aspose.Cells | customize sparkline line weight and series color in Aspose.Cells C#
// Tags: add line sparkline group Aspose.Cells C# | sparkline group configuration Aspose.Cells | customize sparkline high low markers Aspose.Cells | define sparkline source range A1:E1 Aspose.Cells | insert sparkline into specific cell G1 Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The program loads an existing Excel workbook (or creates a sample one if missing), ensures a data row, defines the source range A1:E1, adds a line‑type sparkline to cell G1, customizes high‑point (green) and low‑point (red) markers, sets the series line color to blue with a line weight of 1, and saves the result as output_sparklines.xlsx.
class Program
{
    static void Main(string[] args)
    {
        // Input workbook path – can be supplied as a command‑line argument.
        string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
        string outputPath = "output_sparklines.xlsx";

        // If the input file does not exist, create a simple workbook with sample data.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a sample workbook.");

            Workbook sampleWb = new Workbook();
            Worksheet sampleWs = sampleWb.Worksheets[0];

            // Populate a 4x4 matrix with sample numbers.
            for (int row = 0; row < 4; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    sampleWs.Cells[row, col].PutValue((row + 1) * (col + 1));
                }
            }

            sampleWb.Save(inputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Sample workbook saved as \"{inputPath}\".");
        }

        try
        {
            // Load the existing workbook.
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one row of data for the sparkline.
            if (sheet.Cells.MaxDataRow < 0 || sheet.Cells.MaxDataColumn < 0)
            {
                Console.WriteLine("Worksheet contains no data. Adding a default data row.");
                for (int i = 0; i < 5; i++)
                {
                    sheet.Cells[0, i].PutValue(i + 1);
                }
            }

            // Define the source data range for the sparkline (first row, columns A‑E).
            string dataRange = "A1:E1";

            // Define the cell area where the sparkline will be placed (column G, row 1).
            CellArea sparklineArea = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 6, // Column G (0‑based index)
                EndColumn = 6
            };

            // Add a line‑type sparkline group.
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, dataRange, false, sparklineArea);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Customize the sparkline appearance.
            group.ShowHighPoint = true;
            group.ShowLowPoint = true;
            group.HighPointColor.Color = Color.Green;
            group.LowPointColor.Color = Color.Red;
            group.LineWeight = 1.0;

            // Set the series color (the line itself).
            CellsColor seriesClr = workbook.CreateCellsColor();
            seriesClr.Color = Color.Blue;
            group.SeriesColor = seriesClr;

            // Save the workbook with the newly created sparkline.
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook with sparkline saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while processing the workbook:");
            Console.WriteLine(ex.ToString());
        }
    }
}
