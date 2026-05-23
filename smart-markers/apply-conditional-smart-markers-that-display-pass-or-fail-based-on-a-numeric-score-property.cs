using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Markup;

namespace ConditionalSmartMarkerDemoApp
{
    // Data class containing the numeric score
    public class ScoreData
    {
        public int Score { get; set; }
    }

    // Callback that replaces the "Result" smart marker with "Pass"/"Fail"
    public class ResultSmartMarkerCallback : ISmartMarkerCallBack
    {
        private readonly Workbook _workbook;

        public ResultSmartMarkerCallback(Workbook workbook)
        {
            _workbook = workbook;
        }

        // Called for each smart marker during processing
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Intervene only for the "Result" column
            if (columnName.Equals("Result", StringComparison.OrdinalIgnoreCase))
            {
                // Score is in column A (index 0) of the same row
                Cell scoreCell = _workbook.Worksheets[sheetIndex].Cells[rowIndex, 0];
                object scoreObj = scoreCell.Value;

                // If the score is numeric, write Pass/Fail; otherwise leave blank
                if (scoreObj is int score)
                {
                    string result = score >= 60 ? "Pass" : "Fail";
                    _workbook.Worksheets[sheetIndex].Cells[rowIndex, colIndex].PutValue(result);
                }
            }
        }
    }

    public class ConditionalSmartMarkerDemo
    {
        public static void Run()
        {
            try
            {
                // ---------- Create a new workbook (lifecycle: create) ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Header row
                sheet.Cells["A1"].PutValue("Score");
                sheet.Cells["B1"].PutValue("Result");

                // Row 2 contains smart markers
                // &=$Score will be replaced by the numeric score from the data source
                sheet.Cells["A2"].PutValue("&=$Score");
                // &=$Result will be processed by our callback to show Pass/Fail
                sheet.Cells["B2"].PutValue("&=$Result");

                // ---------- Prepare data source ----------
                var data = new List<ScoreData>
                {
                    new ScoreData { Score = 85 },
                    new ScoreData { Score = 45 },
                    new ScoreData { Score = 73 }
                };

                // ---------- Set up WorkbookDesigner ----------
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook,
                    // Assign the callback that will handle the "Result" smart marker
                    CallBack = new ResultSmartMarkerCallback(workbook)
                };

                // The data source name must match the smart marker prefix (default is empty)
                designer.SetDataSource("Score", data);

                // Process all smart markers (lifecycle: process)
                designer.Process(true);

                // ---------- Save the result (lifecycle: save) ----------
                string outputPath = "ConditionalSmartMarkerResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during demo execution: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                ConditionalSmartMarkerDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}