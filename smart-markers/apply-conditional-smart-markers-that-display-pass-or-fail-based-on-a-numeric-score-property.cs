// Title: C# Conditional Smart Marker in Aspose.Cells – Show Pass/Fail from Score
// Description: The sample creates a workbook, places a smart‑marker formula that evaluates a numeric Score, binds a List<TestResult> as the data source, runs WorkbookDesigner, and saves an Excel file where each row reads “Pass” for scores 60 or higher and “Fail” otherwise.
// Keywords: Aspose.Cells | C# smart marker | conditional expression | pass fail logic | WorkbookDesigner | list data source | Excel automation | score threshold | dynamic cell content | Aspose.Cells .NET
// Common Searches: Aspose.Cells C# smart marker pass fail | how to use conditional expression in Aspose.Cells | binding List<T> to smart markers .NET | C# generate Excel with pass/fail based on score | WorkbookDesigner conditional text example
// Developer Intent: Produce an Excel report that automatically labels each entry as Pass or Fail using a smart‑marker rule.
// Use Cases: Student exam result sheets that automatically display pass‑fail status. | Employee performance dashboards that highlight individuals meeting a target threshold. | Quality‑control logs that flag items falling below a defined score.
// AI Prompts: Create C# code with Aspose.Cells that adds a smart‑marker evaluating a numeric field and writes 'Pass' if the value is ≥ 60, otherwise 'Fail'. | Explain how to bind a List<TestResult> to WorkbookDesigner and process a conditional smart‑marker expression. | Show how to modify the threshold and customize the output messages in the smart‑marker formula.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsConditionalSmartMarkerDemo
{
    // Simple data class with a numeric Score property
    // The sample creates a workbook, places a smart‑marker formula that evaluates a numeric Score, binds a List<TestResult> as the data source, runs WorkbookDesigner, and saves an Excel file where each row reads “Pass” for scores 60 or higher and “Fail” otherwise.
    public class TestResult
    {
        public int Score { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a conditional smart marker that displays "Pass" if Score >= 60, otherwise "Fail"
            // Syntax: &=$Score>=60?"Pass":"Fail"
            sheet.Cells["A1"].PutValue("&=$Score>=60?\"Pass\":\"Fail\"");

            // Prepare sample data source
            List<TestResult> results = new List<TestResult>
            {
                new TestResult { Score = 85 }, // Should display "Pass"
                new TestResult { Score = 45 }  // Should display "Fail"
            };

            // Set up the workbook designer and bind the data source
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("Data", results);

            // Process the smart markers (process rule)
            designer.Process();

            // Save the workbook (save rule)
            workbook.Save("ConditionalSmartMarkerOutput.xlsx");
        }
    }
}
