// Title: C# – Extract Linear Trendline Equation from an Aspose.Cells Chart
// Description: Loads or creates an Excel workbook, ensures a line chart exists, adds a linear trendline with equation display, saves the file, then retrieves the trendline equation using the TrendlineEquation property (or GetTrendlineEquation via reflection for older versions) and writes it to the console.
// Keywords: Aspose.Cells trendline equation | C# extract chart trendline | Aspose.Cells linear trendline | retrieve trendline formula | Aspose.Cells reflection | Excel chart regression line | Aspose.Cells C# example
// Common Searches: how to get trendline equation with Aspose.Cells C# | Aspose.Cells get linear regression formula from chart | extract chart trendline formula using Aspose.Cells | Aspose.Cells TrendlineEquation property missing | C# read trendline equation from Excel chart
// Developer Intent: Add a linear trendline to a chart series and programmatically obtain its equation for further processing.
// Use Cases: Display the regression equation alongside charts in automated reports. | Validate trendline calculations in unit tests by comparing extracted formulas. | Populate dashboard text boxes with the trendline formula for end‑user insight.
// AI Prompts: Generate C# code that adds a polynomial trendline to an Aspose.Cells chart and returns its equation. | Create a reusable method to extract the trendline equation from any chart series, handling both current and legacy Aspose.Cells APIs. | Explain how to format the extracted trendline equation for inclusion in a PDF generated with Aspose.Words.

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads or creates an Excel workbook, ensures a line chart exists, adds a linear trendline with equation display, saves the file, then retrieves the trendline equation using the TrendlineEquation property (or GetTrendlineEquation via reflection for older versions) and writes it to the console.
class TrendlineFormulaExtractor
{
    static void Main()
    {
        try
        {
            // Path to the workbook (replace with your actual file if needed)
            string inputPath = "input.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one chart; if not, create a simple line chart for demonstration
            Chart chart;
            if (worksheet.Charts.Count > 0)
            {
                chart = worksheet.Charts[0];
            }
            else
            {
                // Add sample data for the chart
                worksheet.Cells["A1"].PutValue(1);
                worksheet.Cells["A2"].PutValue(2);
                worksheet.Cells["A3"].PutValue(3);
                worksheet.Cells["A4"].PutValue(4);
                worksheet.Cells["B1"].PutValue(2);
                worksheet.Cells["B2"].PutValue(4);
                worksheet.Cells["B3"].PutValue(6);
                worksheet.Cells["B4"].PutValue(8);

                // Create a line chart
                int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 15, 5);
                chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("A1:B4", true);
            }

            // Add a linear trendline to the first series
            int trendlineIndex = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
            Trendline trendline = chart.NSeries[0].TrendLines[trendlineIndex];
            trendline.DisplayEquation = true; // Enable equation display

            // Save the workbook (required to persist the trendline equation)
            string tempPath = "temp_with_trendline.xlsx";
            workbook.Save(tempPath);

            // Retrieve the trendline equation using reflection (covers different API versions)
            string equation = null;
            try
            {
                PropertyInfo eqProp = typeof(Trendline).GetProperty("TrendlineEquation");
                if (eqProp != null)
                {
                    equation = eqProp.GetValue(trendline) as string;
                }
                else
                {
                    // Fallback to a possible method name in older versions
                    MethodInfo getEqMethod = typeof(Trendline).GetMethod("GetTrendlineEquation");
                    if (getEqMethod != null)
                    {
                        equation = getEqMethod.Invoke(trendline, null) as string;
                    }
                }
            }
            catch (Exception reflEx)
            {
                Console.WriteLine("Reflection error while retrieving equation: " + reflEx.Message);
            }

            if (string.IsNullOrEmpty(equation))
            {
                equation = "Trendline equation not available in this Aspose.Cells version.";
            }

            // Output the extracted equation to the console
            Console.WriteLine("Extracted Trendline Equation:");
            Console.WriteLine(equation);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
