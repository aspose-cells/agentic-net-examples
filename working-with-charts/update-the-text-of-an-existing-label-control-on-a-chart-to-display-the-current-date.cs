// Title: Set a chart label to today's date with Aspose.Cells for .NET
// Description: Loads a workbook, opens the first worksheet and its first chart, locates the first Label shape inside the chart, assigns DateTime.Now (short‑date format) to the label's Text property, and saves the updated file.
// Keywords: Aspose.Cells chart label | C# update chart textbox | set chart title date | modify Excel chart annotation | Aspose.Cells label shape | DateTime.Now chart label | Excel chart dynamic date | Aspose.Cells .NET example
// Common Searches: Aspose.Cells change chart label text | C# set chart label to current date | How to update chart textbox in Excel using Aspose.Cells | Find and edit label shape in Aspose.Cells chart | Add timestamp to Excel chart programmatically
// Developer Intent: Replace the text of an existing chart label with the current date using C# and Aspose.Cells.
// Use Cases: Insert a generation timestamp into a chart title before distributing a daily report. | Automate batch updates of multiple workbooks to show the processing date on each chart annotation. | Swap placeholder text in a chart label with today's date for a rolling dashboard.
// AI Prompts: Write C# code with Aspose.Cells that finds the first Label shape in the first chart of a worksheet and sets its Text to DateTime.Now formatted as a short date. | Show how to loop through all charts in a workbook and update every label shape to display the current date before saving. | Provide an example that checks for the presence of a label shape in a chart, updates its text to today’s date, and gracefully handles the case when no label is found.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Loads a workbook, opens the first worksheet and its first chart, locates the first Label shape inside the chart, assigns DateTime.Now (short‑date format) to the label's Text property, and saves the updated file.
class UpdateChartLabel
{
    static void Main()
    {
        // Load an existing workbook that already contains a chart with a label
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Assume the chart we want to modify is the first chart on the sheet
        if (sheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }
        Chart chart = sheet.Charts[0];

        // Find the first label shape inside the chart.
        // Labels added to a chart are stored in the chart's Shapes collection.
        Label chartLabel = null;
        foreach (Shape shape in chart.Shapes)
        {
            // The Shape type for a label is AutoShapeType.TextBox (or generic Shape with Text property)
            // We treat any shape that has a Text property and is not a data series as a label.
            // Here we simply take the first shape that is a Label.
            if (shape is Label)
            {
                chartLabel = (Label)shape;
                break;
            }
        }

        if (chartLabel == null)
        {
            Console.WriteLine("No label shape found in the chart.");
            return;
        }

        // Update the label text to the current date.
        // Use a short date format; adjust as required.
        chartLabel.Text = DateTime.Now.ToString("d");

        // Save the workbook with the updated label.
        workbook.Save("output.xlsx");
    }
}
