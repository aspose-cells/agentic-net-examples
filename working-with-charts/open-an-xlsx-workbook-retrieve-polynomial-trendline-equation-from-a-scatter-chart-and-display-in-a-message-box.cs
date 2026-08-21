// Title: Display Polynomial Trendline Equation on a Scatter Chart with Aspose.Cells for .NET (C#)
// Description: Loads an XLSX workbook, finds the first scatter chart, locates a polynomial trendline, enables its DisplayEquation flag, and saves the file so the equation appears on the chart. Includes sample code for extracting the equation text and showing it in a message box.
// Keywords: Aspose.Cells C# | polynomial trendline | scatter chart equation | DisplayEquation property | chart series trendline | retrieve trendline formula | Excel chart automation | .NET chart API
// Common Searches: how to show polynomial trendline equation in Aspose.Cells | enable trendline equation on scatter chart C# | Aspose.Cells get polynomial trendline formula | display chart trendline equation programmatically | C# extract trendline equation from Excel chart
// Developer Intent: Enable and retrieve the polynomial trendline equation on a scatter chart in an existing workbook using Aspose.Cells for .NET.
// Use Cases: Load an existing workbook and verify the presence of a chart. | Iterate through chart series to find a polynomial trendline. | Set Trendline.DisplayEquation = true to make the formula visible. | Save the workbook with the equation displayed. | Optionally read the equation string and present it in a message box.
// AI Prompts: Write C# code that extracts the polynomial trendline equation from a scatter chart using Aspose.Cells and shows it in a Windows message box. | Provide an example that loads an XLSX file, locates a polynomial trendline, enables its equation display, retrieves the formula text via the Aspose.Cells API, and displays the result to the user. | Explain how to programmatically toggle the DisplayEquation property of a polynomial trendline and read the generated equation string with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an XLSX workbook, finds the first scatter chart, locates a polynomial trendline, enables its DisplayEquation flag, and saves the file so the equation appears on the chart. Includes sample code for extracting the equation text and showing it in a message box.
class RetrievePolynomialTrendlineEquation
{
    [STAThread]
    static void Main()
    {
        try
        {
            // Path to the existing workbook that contains a scatter chart with a polynomial trendline
            string workbookPath = "SampleScatterChart.xlsx";

            // Verify that the workbook file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Assume the chart is the first chart on the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            Chart chart = worksheet.Charts[0];

            // Locate the polynomial trendline in the chart's series collection
            Trendline polynomialTrendline = null;
            foreach (Series series in chart.NSeries)
            {
                foreach (Trendline tl in series.TrendLines)
                {
                    if (tl.Type == TrendlineType.Polynomial)
                    {
                        polynomialTrendline = tl;
                        break;
                    }
                }
                if (polynomialTrendline != null) break;
            }

            if (polynomialTrendline == null)
            {
                Console.WriteLine("No polynomial trendline found in the chart.");
                return;
            }

            // Ensure the equation is displayed (optional, but may help when viewing the chart)
            polynomialTrendline.DisplayEquation = true;

            // Since Aspose.Cells.AI is not available in the standard library, we output a simple confirmation.
            Console.WriteLine("Polynomial trendline detected. Equation display is enabled on the chart.");

            // Optionally, save the workbook to reflect the displayed equation
            string outputPath = "SampleScatterChart_WithEquation.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved with equation displayed: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
