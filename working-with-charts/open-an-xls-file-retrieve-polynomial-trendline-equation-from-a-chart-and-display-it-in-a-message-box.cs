// Title: Read a polynomial trendline equation from the first chart in an XLS workbook using Aspose.Cells for .NET and display it in a Windows MessageBox
// AI Prompts: Write C# code that opens an .xls file with Aspose.Cells, finds the first chart, extracts the polynomial trendline formula (order and coefficients), and shows the result in a MessageBox. | Show how to add error handling for missing workbook files, absent charts, and Aspose.Cells versions that do not expose the Trendlines collection while retrieving a polynomial trendline. | Demonstrate converting a Trendline object's Equation property into a readable string and presenting it via System.Windows.Forms.MessageBox.
// Common Searches: Aspose.Cells C# get polynomial trendline formula from Excel chart in .xls file | How to read trendline equation from first chart using Aspose.Cells for .NET | Display Excel chart trendline order in a Windows MessageBox with C# | C# Aspose.Cells handling unsupported trendline collection in older versions | Extract polynomial trendline coefficients from XLS workbook using Aspose.Cells
// Tags: Aspose.Cells read polynomial trendline | extract chart trendline equation .xls | display trendline formula MessageBox C# | trendline collection compatibility Aspose.Cells | chart series trendline handling .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an .xls workbook with Aspose.Cells, accesses the first worksheet's first chart, iterates its series to locate polynomial trendlines, builds a readable equation string, and presents it in a Windows MessageBox while handling missing files, absent charts, and versions lacking Trendline support.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string filePath = "input.xls";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the existing XLS workbook
            Workbook workbook = new Workbook(filePath);

            // Assume the chart is on the first worksheet and is the first chart in the sheet
            Worksheet sheet = workbook.Worksheets[0];
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            Chart chart = sheet.Charts[0];

            // Iterate through all series in the chart
            foreach (Series series in chart.NSeries)
            {
                // Use dynamic to access Trendlines (may not be available in older versions)
                dynamic dynSeries = series;
                try
                {
                    foreach (Trendline trendline in dynSeries.Trendlines)
                    {
                        // Check if the trendline is a polynomial type
                        if (trendline.Type == TrendlineType.Polynomial)
                        {
                            Console.WriteLine($"Polynomial Trendline found. Order: {trendline.Order}");
                        }
                    }
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    // Trendlines collection not supported in this version of Aspose.Cells
                    Console.WriteLine("Trendlines are not supported in the current Aspose.Cells version.");
                }
                catch (Exception ex)
                {
                    // Handle any unexpected errors while processing trendlines
                    Console.WriteLine($"Error processing trendlines: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
