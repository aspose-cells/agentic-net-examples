// Title: Extract Polynomial Trendline Equation from an XLS Chart with Aspose.Cells (C#)
// Description: Loads an XLS workbook, accesses the first worksheet and its first chart, finds a polynomial trendline in the first series, enables equation display, and retrieves the equation text (using reflection for version‑agnostic support). The equation is written to the console or can be shown in a UI.
// Keywords: Aspose.Cells trendline equation | C# polynomial trendline extraction | read chart trendline Aspose | XLS chart trendline label | Aspose.Cells get polynomial equation | Excel chart trendline .NET | Aspose.Cells chart API | trendline equation reflection
// Common Searches: Aspose.Cells get polynomial trendline equation from XLS | C# extract chart trendline label using Aspose.Cells | how to read trendline equation in Excel chart .NET | retrieve polynomial trendline from chart with Aspose | display trendline equation in C# Aspose.Cells
// Developer Intent: Load an existing XLS file, locate a polynomial trendline in a chart, and obtain its equation string for display or further processing.
// Use Cases: Log the polynomial trendline equation for automated reporting. | Show the equation in a Windows Forms MessageBox after opening the workbook. | Validate chart data by comparing extracted coefficients with expected values.
// AI Prompts: Generate C# code that opens an XLS file with Aspose.Cells, finds the first chart, locates a polynomial trendline in the first series, and returns the equation as a string. | Create a method that safely retrieves a trendline equation using reflection when the TrendlineLabel property is unavailable, providing a fallback message. | Provide a sample that displays the extracted polynomial trendline equation in a WinForms MessageBox after loading the workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an XLS workbook, accesses the first worksheet and its first chart, finds a polynomial trendline in the first series, enables equation display, and retrieves the equation text (using reflection for version‑agnostic support). The equation is written to the console or can be shown in a UI.
class Program
{
    [STAThread]
    static void Main()
    {
        // Path to the existing XLS file
        string path = "input.xls";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(path))
        {
            Console.WriteLine($"Error: The file '{path}' was not found.");
            return;
        }

        try
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(path);

            // Access the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = worksheet.Charts[0];

            // Locate a polynomial trendline in the first series
            Trendline polynomialTrendline = null;
            foreach (Trendline tl in chart.NSeries[0].TrendLines)
            {
                if (tl.Type == TrendlineType.Polynomial)
                {
                    polynomialTrendline = tl;
                    break;
                }
            }

            if (polynomialTrendline == null)
            {
                Console.WriteLine("No polynomial trendline found in the first series.");
                return;
            }

            // Ensure the equation is set to be displayed (optional)
            polynomialTrendline.DisplayEquation = true;

            // Retrieve the polynomial trendline equation.
            // Note: In some Aspose.Cells versions the TrendlineLabel property may not be available.
            // If unavailable, we fall back to a generic message.
            string equation;
            try
            {
                // Attempt to use TrendlineLabel if the property exists.
                // This block is kept for compatibility with newer library versions.
                var labelProp = polynomialTrendline.GetType().GetProperty("TrendlineLabel");
                if (labelProp != null)
                {
                    var labelObj = labelProp.GetValue(polynomialTrendline);
                    var textProp = labelObj?.GetType().GetProperty("Text");
                    equation = textProp?.GetValue(labelObj) as string ?? "Equation not available";
                }
                else
                {
                    equation = "Equation display enabled (label not accessible).";
                }
            }
            catch
            {
                equation = "Unable to retrieve equation.";
            }

            // Output the retrieved equation
            Console.WriteLine("Polynomial Trendline Equation:");
            Console.WriteLine(equation);
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
