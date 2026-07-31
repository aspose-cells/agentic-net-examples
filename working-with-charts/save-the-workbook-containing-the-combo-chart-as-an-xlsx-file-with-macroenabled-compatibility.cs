// Title: Aspose.Cells C# – Create a column‑line combo chart and export to macro‑enabled XLSM
// Description: A C# sample that builds a workbook, fills cells A1:C4 with data, adds a column‑line combo chart, sets series names and category axis, then saves the file as a macro‑enabled XLSM using Aspose.Cells.
// Keywords: Aspose.Cells | C# combo chart | macro enabled XLSM | SaveFormat.Xlsm | Excel chart automation | column line combo | Aspose.Cells chart example | export chart to XLSM | create workbook with chart | Aspose.Cells SaveFormat
// Common Searches: Aspose.Cells create combo chart C# | save workbook as XLSM Aspose.Cells | C# export chart to macro enabled file | add line series to column chart Aspose.Cells | Aspose.Cells SaveFormat.Xlsm example
// Developer Intent: Generate a workbook containing a column‑line combo chart and write it to a macro‑enabled XLSM file.
// Use Cases: Automate monthly sales reports that include a combo chart and need to be distributed as macro‑enabled files for further user interaction. | Create Excel templates with pre‑defined charts and VBA macros, populate them programmatically, and save in XLSM format. | Build interactive dashboards that combine visual charts with macro functionality, exporting the result as a macro‑enabled workbook.
// AI Prompts: Show how to add a chart title and axis labels to the combo chart before saving the workbook as XLSM. | Provide code to embed a simple VBA macro into the workbook after creating the combo chart and then save it as a macro‑enabled file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboChartMacro
{
    // A C# sample that builds a workbook, fills cells A1:C4 with data, adds a column‑line combo chart, sets series names and category axis, then saves the file as a macro‑enabled XLSM using Aspose.Cells.
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
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Series 2");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(20);

            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["C3"].PutValue(25);

            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(20);
            sheet.Cells["C4"].PutValue(30);

            // Add a combo chart (Column + Line)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // First series as Column
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries[0].Name = "Series 1";

            // Second series as Line
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries[1].Name = "Series 2";
            chart.NSeries[1].Type = ChartType.Line;

            // Set category (X) axis data
            chart.NSeries.CategoryData = "A2:A4";

            // Save the workbook as a macro‑enabled XLSM file
            workbook.Save("ComboChart.xlsm", SaveFormat.Xlsm);
        }
    }
}
