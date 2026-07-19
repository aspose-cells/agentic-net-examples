// Title: Save a Combo Chart Workbook as a Macro‑Enabled XLSM File with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a column‑line combo chart, enable macros, and export the file as a macro‑enabled XLSM using Aspose.Cells SaveFormat.Xlsm in C#.
// Keywords: Aspose.Cells | C# | .NET | combo chart | column chart | line chart | macro enabled workbook | XLSM export | SaveFormat.Xlsm | enable macros | chart series type | Excel automation
// Common Searches: Aspose.Cells save workbook as xlsm C# | create combo chart Aspose.Cells .NET | enable macros in Aspose.Cells workbook | export Excel file with macros using Aspose.Cells | how to save macro‑enabled file with Aspose.Cells
// Developer Intent: Export a workbook that contains a combo chart to a macro‑enabled XLSM file using Aspose.Cells for .NET.
// Use Cases: Generate sales or financial reports with column‑line combo charts and distribute them as macro‑enabled templates. | Create Excel dashboards that require VBA automation alongside visual charts, then save as .xlsm for end‑users. | Automate the production of Excel files that combine chart visualizations with embedded macros for downstream processing.
// AI Prompts: Write C# code with Aspose.Cells to build a column‑line combo chart and save the workbook as a macro‑enabled XLSM file. | Explain how to enable macros in an Aspose.Cells workbook and which SaveFormat should be used for macro‑enabled files. | Show how to change the chart type of a series in an Aspose.Cells combo chart before exporting to XLSM.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add a column‑line combo chart, enable macros, and export the file as a macro‑enabled XLSM using Aspose.Cells SaveFormat.Xlsm in C#.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the combo chart
        worksheet.Cells["A1"].PutValue("Month");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["C1"].PutValue("Profit");

        worksheet.Cells["A2"].PutValue("Jan");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["C2"].PutValue(30);

        worksheet.Cells["A3"].PutValue("Feb");
        worksheet.Cells["B3"].PutValue(150);
        worksheet.Cells["C3"].PutValue(45);

        worksheet.Cells["A4"].PutValue("Mar");
        worksheet.Cells["B4"].PutValue(180);
        worksheet.Cells["C4"].PutValue(60);

        // Add a combo chart: Sales as Column, Profit as Line
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
        Chart comboChart = worksheet.Charts[chartIndex];

        // First series (Column)
        comboChart.NSeries.Add("B2:B4", true);
        comboChart.NSeries[0].Type = ChartType.Column;

        // Second series (Line)
        comboChart.NSeries.Add("C2:C4", true);
        comboChart.NSeries[1].Type = ChartType.Line;

        // Set category (X) axis data
        comboChart.NSeries.CategoryData = "A2:A4";

        // Enable macros in the workbook (optional but aligns with macro‑enabled format)
        workbook.Settings.EnableMacros = true;

        // Save the workbook as a macro‑enabled file (Xlsm) using the Save(string, SaveFormat) rule
        workbook.Save("ComboChartMacroEnabled.xlsm", SaveFormat.Xlsm);
    }
}
