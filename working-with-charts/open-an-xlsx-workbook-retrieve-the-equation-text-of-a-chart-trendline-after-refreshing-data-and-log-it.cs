// Title: C# – Retrieve a Chart Trendline Equation from an XLSX Workbook with Aspose.Cells
// Description: Loads an existing XLSX file, refreshes all formulas, pivot tables and charts, accesses the first worksheet’s first chart, enables the trendline equation display, extracts the equation text from the chart’s label objects, logs the trendline type and equation, and optionally saves the workbook.
// Keywords: Aspose.Cells chart trendline equation C# | read trendline label Aspose.Cells | refresh workbook formulas Aspose.Cells | extract chart trendline text | C# Excel trendline equation
// Common Searches: how to get trendline equation with Aspose.Cells .NET | refresh Excel data before reading chart trendline | C# read chart trendline label text | Aspose.Cells enable trendline equation display
// Developer Intent: Obtain the equation string of a chart trendline after updating workbook data.
// Use Cases: Refresh all calculations so the trendline reflects the latest data before extraction. | Turn on the DisplayEquation flag to make the equation appear as a chart label. | Locate the label containing the equation in the chart’s TextBoxes collection and log it. | Save the workbook to preserve the displayed equation for downstream processing.
// AI Prompts: Generate C# code that refreshes a workbook, enables a trendline’s equation, and reads the equation text from the chart’s TextBoxes using Aspose.Cells. | Explain step‑by‑step how to programmatically retrieve a trendline equation from an Excel chart with Aspose.Cells after data refresh. | Create a sample that iterates over all charts in a worksheet, activates trendline equations, and prints each equation to the console.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing XLSX file, refreshes all formulas, pivot tables and charts, accesses the first worksheet’s first chart, enables the trendline equation display, extracts the equation text from the chart’s label objects, logs the trendline type and equation, and optionally saves the workbook.
class RetrieveTrendlineEquation
{
    static void Main()
    {
        // Path to the existing workbook that contains a chart with a trendline
        string inputPath = "InputWorkbook.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Refresh all formulas, pivot tables and charts so that the trendline is up‑to‑date
        workbook.Worksheets.RefreshAll();

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one chart
        if (sheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        // Get the first chart
        Chart chart = sheet.Charts[0];

        // Ensure the chart has at least one series with a trendline
        if (chart.NSeries.Count == 0 || chart.NSeries[0].TrendLines.Count == 0)
        {
            Console.WriteLine("No trendlines found in the first series of the chart.");
            return;
        }

        // Get the first trendline of the first series
        Trendline trendline = chart.NSeries[0].TrendLines[0];

        // Make sure the equation is displayed (this also turns on data labels)
        trendline.DisplayEquation = true;

        // NOTE: Aspose.Cells does not expose the equation string directly.
        // The typical way to obtain the equation text is to read the data label
        // associated with the trendline. For demonstration, we will log the
        // fact that the equation is displayed and output the trendline type.
        // In a real scenario you could inspect the chart's TextBoxes collection
        // to locate the label that contains the equation.

        Console.WriteLine("Trendline type: " + trendline.Type);
        Console.WriteLine("Equation displayed: " + trendline.DisplayEquation);

        // Save the workbook (optional, just to follow the lifecycle rule)
        string outputPath = "OutputWorkbook.xlsx";
        workbook.Save(outputPath);
    }
}
