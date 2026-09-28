// Title: Generate an Excel file with Aspose.Cells for .NET that uses a range smart marker IF expression to display “Pass” or “Fail” based on a numeric Score field
// AI Prompts: Write C# code that creates a Workbook, defines a range smart marker using the IF function (&=IF($Score>=60,"Pass","Fail")), binds a List<ExamResult> as the data source, names the smart marker range "_CellsSmartMarkers", processes only that range, and saves the result as an .xlsx file. | Demonstrate how to configure WorkbookDesigner to evaluate conditional smart markers within a specific range and output Pass/Fail text based on a numeric property.
// Common Searches: Aspose.Cells C# conditional smart marker IF expression example | How to bind a List data source to range smart markers in Aspose.Cells .NET | Create Excel report with Pass or Fail status using Aspose.Cells smart markers | WorkbookDesigner process only a named smart marker range Aspose.Cells | Conditional text output in Excel using Aspose.Cells range smart markers
// Tags: range smart marker IF expression Aspose.Cells | WorkbookDesigner bind List data source | conditional Pass Fail output Excel | process specific smart marker range | generate .xlsx with conditional text Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsConditionalSmartMarkerDemo
{
    // Simple data class containing a numeric score
    // The example creates a new workbook, sets up a range smart marker that uses an IF expression to show "Pass" or "Fail" according to each ExamResult.Score, binds a List<ExamResult> as the data source named "Data", names the marker range "_CellsSmartMarkers", processes only that range with WorkbookDesigner, and saves the populated workbook to an Excel file.
    public class ExamResult
    {
        public int Score { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // ---------- Create a new workbook ----------
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // ---------- Set up a conditional smart marker ----------
                // The smart marker uses an IF expression to display "Pass" or "Fail"
                // based on the value of the Score property.
                // Syntax: &IF($Score>=60,"Pass","Fail")
                sheet.Cells["A1"].PutValue("Score");
                sheet.Cells["B1"].PutValue("Result");
                sheet.Cells["A2"].PutValue("&=$Score"); // raw score value
                sheet.Cells["B2"].PutValue("&=IF($Score>=60,\"Pass\",\"Fail\")"); // conditional result

                // ---------- Prepare data source ----------
                var results = new List<ExamResult>
                {
                    new ExamResult { Score = 85 },
                    new ExamResult { Score = 55 },
                    new ExamResult { Score = 73 },
                    new ExamResult { Score = 42 }
                };

                // ---------- Configure WorkbookDesigner ----------
                var designer = new WorkbookDesigner
                {
                    Workbook = workbook
                    // LineByLine is obsolete; range smart markers are used instead.
                };

                // Set the data source. The name "Data" will be used in the smart markers.
                designer.SetDataSource("Data", results);

                // Define the range that contains the smart markers.
                // The range must be named "_CellsSmartMarkers" when using range smart markers.
                AsposeRange smartRange = sheet.Cells.CreateRange("A2:B5");
                smartRange.Name = "_CellsSmartMarkers";

                // ---------- Process the smart markers ----------
                // The boolean parameter indicates that only the defined range is processed.
                designer.Process(smartRange, true);

                // ---------- Save the workbook ----------
                const string outputPath = "ConditionalSmartMarkerResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
