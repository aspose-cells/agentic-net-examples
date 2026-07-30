// Title: C# – Get Polynomial Trendline Equation from First Chart in XLS using Aspose.Cells
// Description: Loads an .xls workbook with Aspose.Cells, accesses the first worksheet and its first chart, scans the series for a polynomial trendline, and returns the equation (or formula) for display or logging.
// Keywords: Aspose.Cells polynomial trendline | C# get trendline equation | extract chart trendline .xls | Aspose.Cells chart API | reflection trendline property | Excel polynomial trendline equation | Aspose.Cells GetTrendlineEquation | C# Excel chart trendline | Aspose.Cells .NET chart trendline | retrieve polynomial trendline formula
// Common Searches: How to extract polynomial trendline equation from an Excel chart using Aspose.Cells C# | C# Aspose.Cells get trendline formula from .xls chart | Read polynomial trendline equation with Aspose.Cells .NET | Aspose.Cells chart trendline extraction example | Retrieve polynomial trendline equation from first chart in workbook
// Developer Intent: Extract and display the polynomial trendline equation from the first chart in an XLS workbook using Aspose.Cells for .NET.
// Use Cases: Automate reporting by pulling the polynomial trendline equation from existing Excel files. | Integrate equation extraction into a Windows Forms or WPF application for user‑visible results. | Use reflection to maintain compatibility across different Aspose.Cells library versions. | Process multiple workbooks to collect trendline equations for data analysis.
// AI Prompts: Generate C# code that opens an .xls file with Aspose.Cells, finds the first chart, and returns the polynomial trendline equation, using reflection to handle missing properties. | Show how to display the extracted polynomial trendline equation in a MessageBox after retrieving it from a chart. | Explain how to extend the method to handle several charts and select a specific series by its name or index. | Provide a PowerShell script that calls the compiled .NET assembly to output the trendline equation from a given workbook.

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Loads an .xls workbook with Aspose.Cells, accesses the first worksheet and its first chart, scans the series for a polynomial trendline, and returns the equation (or formula) for display or logging.
    class Program
    {
        static void Main()
        {
            try
            {
                string filePath = "sample.xls";

                // Verify that the input file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Get the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Assume the chart of interest is the first chart
                Chart chart = sheet.Charts[0];

                // Retrieve polynomial trendline equation
                string equation = GetPolynomialTrendlineEquation(chart);

                if (string.IsNullOrEmpty(equation))
                {
                    Console.WriteLine("No polynomial trendline found on the chart.");
                }
                else
                {
                    Console.WriteLine("Polynomial Trendline Equation:");
                    Console.WriteLine(equation);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Returns the equation of the first polynomial trendline in the specified chart
        private static string GetPolynomialTrendlineEquation(Chart chart)
        {
            try
            {
                foreach (Series series in chart.NSeries)
                {
                    // Use reflection to access Trendlines collection (handles versions where the property may be missing)
                    var trendlinesProp = series.GetType().GetProperty("Trendlines");
                    if (trendlinesProp == null) continue;

                    var trendlinesObj = trendlinesProp.GetValue(series) as IEnumerable;
                    if (trendlinesObj == null) continue;

                    foreach (object trendline in trendlinesObj)
                    {
                        // Access Trendline.Type
                        var typeProp = trendline.GetType().GetProperty("Type");
                        if (typeProp == null) continue;

                        var typeValue = typeProp.GetValue(trendline);
                        if (typeValue == null) continue;

                        if ((TrendlineType)typeValue == TrendlineType.Polynomial)
                        {
                            // Attempt to read the Equation property (may be named differently in some versions)
                            var equationProp = trendline.GetType().GetProperty("Equation");
                            if (equationProp != null)
                            {
                                var eq = equationProp.GetValue(trendline) as string;
                                if (!string.IsNullOrEmpty(eq))
                                    return eq;
                            }

                            // Fallback: some versions expose the formula via the "Formula" property
                            var formulaProp = trendline.GetType().GetProperty("Formula");
                            if (formulaProp != null)
                            {
                                var eq = formulaProp.GetValue(trendline) as string;
                                if (!string.IsNullOrEmpty(eq))
                                    return eq;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while retrieving trendline equation: {ex.Message}");
            }

            return null;
        }
    }
}
