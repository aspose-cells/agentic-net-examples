using System;
using Aspose.Cells;
using System.Drawing;

class WorkbookProcessor
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (or modify as needed)
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Merge cells V8:W9
        // Column V = 21 (zero‑based), Column W = 22
        // Row 8 = index 7, Row 9 = index 8
        cells.Merge(firstRow: 7, firstColumn: 21, totalRows: 2, totalColumns: 2);

        // Apply conditional formatting for negative values on the used range
        // Define the area to which the formatting will be applied (e.g., A1:Z100)
        CellArea formatArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = cells.MaxDataRow,
            EndColumn = cells.MaxDataColumn
        };

        // Add a new ConditionalFormatting collection to the worksheet
        int cfIndex = sheet.ConditionalFormattings.Add();
        FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

        // Associate the defined area with the conditional formatting
        fcc.AddArea(formatArea);

        // Add a condition: cell value less than 0 (negative numbers)
        int conditionIndex = fcc.AddCondition(
            type: FormatConditionType.CellValue,
            operatorType: OperatorType.LessThan,
            formula1: "0",
            formula2: null);

        // Retrieve the created condition and set its style (e.g., red background)
        FormatCondition condition = fcc[conditionIndex];
        Style style = workbook.CreateStyle();
        style.ForegroundColor = Color.Red;
        style.Pattern = BackgroundType.Solid;
        condition.Style = style;

        // Prepare HTML save options (merge conditional formatting areas for compact HTML)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            MergeAreas = true,                     // Merge CF and validation areas
            ExportFormula = true,                  // Keep formulas in HTML (optional)
            ExportActiveWorksheetOnly = true       // Export only the active sheet
        };

        // Save the workbook as HTML
        string outputPath = "output.html";
        workbook.Save(outputPath, htmlOptions);

        Console.WriteLine("Workbook processed and saved as HTML: " + outputPath);
    }
}