// Title: C# – Insert a Stacked Bar Chart into an Existing XLSX Workbook with Aspose.Cells
// Description: Load an XLSX file, create a BarStacked chart on the first worksheet, bind it to a cell range, optionally assign a title, and write the updated workbook to a new file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart automation | stacked bar chart | BarStacked | load XLSX workbook | add chart programmatically | set chart data source | save Excel file | Excel .NET library | Aspose.Cells example
// Common Searches: how to add a stacked bar chart with Aspose.Cells C# | Aspose.Cells load existing workbook and insert chart | C# code for BarStacked chart in Excel file | save workbook after creating chart Aspose.Cells | set chart title Aspose.Cells C#
// Developer Intent: Programmatically place a BarStacked chart into a loaded Excel workbook and persist the modification.
// Use Cases: Produce a quarterly‑sales visual by overlaying a stacked bar chart on a template report. | Automate monthly KPI dashboards that inject bar‑stacked graphics into pre‑filled worksheets. | Generate multi‑sheet financial summaries where each sheet receives its own stacked bar representation.
// AI Prompts: Write C# code with Aspose.Cells to open 'report.xlsx', add a BarStacked chart covering A1:C10 on sheet 1, set the title to 'Quarterly Sales', and save as 'report_with_chart.xlsx'. | Explain how to modify the position and dimensions of a BarStacked chart after it has been added using Aspose.Cells in C#. | Show how to bind a stacked bar chart to a dynamic range that expands with new rows, using Aspose.Cells NSeries.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Load an XLSX file, create a BarStacked chart on the first worksheet, bind it to a cell range, optionally assign a title, and write the updated workbook to a new file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);   // uses Workbook(string) constructor

        // Access the first worksheet (you can change the index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Add a stacked bar chart to the worksheet
        // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
        int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 1, 20, 6);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the chart
        // Adjust the range according to the actual data in the workbook
        chart.NSeries.Add("=Sheet1!$A$1:$B$5", true);

        // Optional: set a title for the chart
        chart.Title.Text = "Stacked Bar Chart";

        // Save the modified workbook
        string outputFile = "output.xlsx";
        workbook.Save(outputFile);   // uses Workbook.Save(string) method
    }
}
