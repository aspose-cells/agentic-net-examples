using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Equations;

class RetrieveTrendlineEquation
{
    static void Main()
    {
        // Load the existing workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Refresh all formulas, pivot tables and charts
        workbook.Worksheets.RefreshAll();

        // Assume the first worksheet and the first chart contain the trendline
        Worksheet sheet = workbook.Worksheets[0];
        if (sheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        Chart chart = sheet.Charts[0];

        // Ensure the first series has at least one trendline and that the equation is displayed
        if (chart.NSeries.Count == 0 || chart.NSeries[0].TrendLines.Count == 0)
        {
            Console.WriteLine("No trendlines found in the first series.");
            return;
        }

        Trendline trendline = chart.NSeries[0].TrendLines[0];
        trendline.DisplayEquation = true; // make sure the equation label is generated

        // After refreshing, the equation is rendered as a TextBox shape inside the chart.
        // Search for a TextBox that contains an equation paragraph.
        string equationText = null;
        foreach (Aspose.Cells.Drawing.Shape shape in chart.Shapes)
        {
            if (shape is TextBox textBox)
            {
                // Try to get the first equation paragraph from the TextBox
                EquationNode eqNode = textBox.GetEquationParagraph();
                if (eqNode != null)
                {
                    // Use LaTeX representation as the equation text
                    equationText = eqNode.ToLaTeX();
                    break;
                }
            }
        }

        if (equationText != null)
        {
            Console.WriteLine("Trendline equation: " + equationText);
        }
        else
        {
            Console.WriteLine("Equation paragraph not found in chart shapes.");
        }

        // Save the workbook (optional, demonstrates the save rule)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}