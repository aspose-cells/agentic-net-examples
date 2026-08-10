// Title: Save a Combo Column‑Line Chart as a Macro‑Enabled XLSM Workbook using Aspose.Cells for .NET (C#)
// Description: Shows how to create a Workbook, fill it with sample data, add a column‑line combo chart, enable macros, and save the result as an XLSM macro‑enabled file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | macro enabled workbook | XLSM | combo chart | column chart | line chart | save as XLSM | enable macros | Excel VBA preservation
// Common Searches: Aspose.Cells save workbook as XLSM | Create combo chart with Aspose.Cells C# | Enable macros in Aspose.Cells workbook | Export chart to macro‑enabled Excel file | C# Aspose.Cells macro enabled file
// Developer Intent: Export a workbook that contains a combo chart to a macro‑enabled XLSM file.
// Use Cases: Generate a sales dashboard with a column‑line chart and deliver it as an XLSM file for downstream VBA processing. | Automate report creation where charts must coexist with embedded macros that need to remain functional after saving. | Copy worksheets that include macros and ensure the macros are preserved when the workbook is saved as XLSM.
// AI Prompts: Write C# code using Aspose.Cells to add a column‑line combo chart and save the workbook as a macro‑enabled XLSM file. | Explain why enabling macros in Aspose.Cells workbook settings is required before saving as XLSM. | Show how to convert a macro‑enabled XLSM workbook to a regular XLSX file while keeping the chart but removing all macros.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboChartMacroSave
{
    // Shows how to create a Workbook, fill it with sample data, add a column‑line combo chart, enable macros, and save the result as an XLSM macro‑enabled file with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the combo chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Column Series");
            sheet.Cells["C1"].PutValue("Line Series");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["C2"].PutValue(30);

            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["C3"].PutValue(45);

            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["C4"].PutValue(60);

            // Add a combo chart (Column + Line)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // First series as Column
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries[0].Name = "Column Series";

            // Second series as Line
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries[1].Name = "Line Series";
            chart.NSeries[1].Type = ChartType.Line;

            // Set category (X) axis data
            chart.NSeries.CategoryData = "A2:A4";

            // Enable macros in workbook settings (optional, required when copying worksheets with macros)
            workbook.Settings.EnableMacros = true;

            // Save the workbook as a macro‑enabled XLSM file
            workbook.Save("ComboChartMacroEnabled.xlsm", SaveFormat.Xlsm);
        }
    }
}
