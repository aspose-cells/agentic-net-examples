// Title: How to set a default row height and auto‑fit only rows that contain formulas using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing workbook (or create a new one), set worksheet.Cells.StandardHeight to 15 points, iterate through each row, use IsFormula to find rows with formulas, call worksheet.AutoFitRow for those rows, and save the file as Xlsx with Aspose.Cells. | Write C# code that opens an Excel file, applies a default row height, identifies rows containing at least one formula cell, auto‑fits only those rows, and writes the result to a new workbook using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# set default row height and auto‑fit rows with formulas | auto fit rows that contain formulas only Aspose.Cells .NET example | detect formula cells and auto‑fit specific rows in a workbook using Aspose.Cells | how to apply StandardHeight then auto‑fit selected rows in Aspose.Cells for C#
// Tags: Aspose.Cells set StandardHeight property | Aspose.Cells AutoFitRow for formula rows | C# iterate worksheet rows to detect formulas | Aspose.Cells save workbook as Xlsx | default row height Excel Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example loads an existing Excel workbook (or creates a new one), sets the worksheet's default row height to 15 points, scans each row for any formula cells, auto‑fits only those rows, and saves the modified workbook as output.xlsx using Aspose.Cells for .NET.
    public class AutoFitRowsWithFormulasDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            string inputPath = "input.xlsx";
            Workbook workbook;

            // Load workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                // Add a sample formula to demonstrate auto‑fit
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].Formula = "=SUM(1,2)";
                ws.Cells["A2"].Value = "Sample text";
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Set default row height (points)
            worksheet.Cells.StandardHeight = 15;

            int maxRow = worksheet.Cells.MaxDataRow;
            int maxCol = worksheet.Cells.MaxDataColumn;

            // Auto‑fit rows that contain at least one formula
            for (int row = 0; row <= maxRow; row++)
            {
                bool containsFormula = false;
                for (int col = 0; col <= maxCol; col++)
                {
                    if (worksheet.Cells[row, col].IsFormula)
                    {
                        containsFormula = true;
                        break;
                    }
                }

                if (containsFormula)
                {
                    worksheet.AutoFitRow(row);
                }
            }

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
