using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportWorksheetPaperDimensionsToCsv
    {
        public static void Run()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "paper_dimensions.csv";

                // Ensure the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(inputPath);

                // Create a new workbook for CSV output
                Workbook csvWorkbook = new Workbook();
                Worksheet csvSheet = csvWorkbook.Worksheets[0];

                // Write header row
                csvSheet.Cells["A1"].PutValue("Worksheet");
                csvSheet.Cells["B1"].PutValue("PaperWidthInches");
                csvSheet.Cells["C1"].PutValue("PaperHeightInches");

                // Export paper dimensions for each worksheet
                for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
                {
                    Worksheet ws = sourceWorkbook.Worksheets[i];
                    double widthInches = ws.PageSetup.PaperWidth / 72.0;   // points to inches
                    double heightInches = ws.PageSetup.PaperHeight / 72.0; // points to inches

                    int rowIndex = i + 2; // data starts at row 2
                    csvSheet.Cells[rowIndex, 0].PutValue(ws.Name);
                    csvSheet.Cells[rowIndex, 1].PutValue(Math.Round(widthInches, 2));
                    csvSheet.Cells[rowIndex, 2].PutValue(Math.Round(heightInches, 2));
                }

                // Save the CSV file
                csvWorkbook.Save(outputPath, SaveFormat.Csv);
                Console.WriteLine($"CSV file saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportWorksheetPaperDimensionsToCsv.Run();
        }
    }
}