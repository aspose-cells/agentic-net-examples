using System;
using Aspose.Cells;

namespace AsposeCellsConditionalFormattingHtml
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A (A1:A5)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["A4"].PutValue(40);
            sheet.Cells["A5"].PutValue(50);

            // Add a new conditional formatting collection
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

            // Define the range that both rules will apply to (A1:A5)
            CellArea area = new CellArea
            {
                StartRow = 0,   // Row 1 (zero‑based)
                EndRow = 4,     // Row 5
                StartColumn = 0,// Column A
                EndColumn = 0   // Column A
            };
            cfCollection.AddArea(area);

            // First rule: cells with value > 25 → red background
            int condIdx1 = cfCollection.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "25",
                null);
            FormatCondition cond1 = cfCollection[condIdx1];
            cond1.Style.BackgroundColor = System.Drawing.Color.Red;
            cond1.StopIfTrue = false; // allow lower‑priority rules to be evaluated

            // Second rule: cells with value < 35 → green background
            int condIdx2 = cfCollection.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.LessThan,
                "35",
                null);
            FormatCondition cond2 = cfCollection[condIdx2];
            cond2.Style.BackgroundColor = System.Drawing.Color.Green;
            cond2.StopIfTrue = false; // ensure both rules can coexist

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export CSS for each worksheet separately so that conditional styles are preserved
                ExportWorksheetCSSSeparately = true,
                // Keep all generated styles (including overlapping ones) in the output
                ExcludeUnusedStyles = false,
                // Export the whole workbook (not only the active sheet)
                ExportActiveWorksheetOnly = false
            };

            // Save the workbook as HTML
            string outputPath = "ConditionalFormattingOverlap.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook saved to HTML at: {outputPath}");
        }
    }
}