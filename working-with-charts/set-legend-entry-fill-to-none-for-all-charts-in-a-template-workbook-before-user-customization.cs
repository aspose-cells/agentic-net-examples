// Title: Set chart legend fill to none for all charts in an Excel template workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, iterates through every worksheet and chart, and sets each chart's legend area FillFormat.FillType to None. | Create a reusable method that accepts a template file path, clears the background fill of all chart legends, and saves the modified workbook to a specified output location while handling missing‑file errors. | Show how to safely access a chart's Legend.Area.FillFormat and apply a transparent fill without affecting other chart properties in Aspose.Cells.
// Common Searches: aspnet aspocells remove chart legend background from all worksheets | c# set legend filltype none for every chart in an existing Excel file | how to make chart legend transparent in Aspose.Cells workbook | bulk update chart legends to no fill in Excel template using Aspose.Cells | iterate charts in Aspose.Cells and clear legend area fill
// Tags: clear legend area fill Aspose.Cells | bulk update chart legends Excel .NET | iterate workbook charts Aspose.Cells C# | transparent chart legend area Aspose.Cells | template workbook legend modification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads a template Excel file with Aspose.Cells, walks through each worksheet and every chart it contains, and sets the legend's area FillFormat.FillType to None, making the legend background transparent. After processing all charts, the workbook is saved to a new file, with error handling for missing templates and other exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            try
            {
                // Ensure the template file exists before loading
                if (!File.Exists(templatePath))
                {
                    throw new FileNotFoundException($"The template file '{templatePath}' was not found.");
                }

                // Load the template workbook
                Workbook workbook = new Workbook(templatePath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts on the worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        // Access the chart's legend
                        Legend legend = chart.Legend;

                        // Set the legend background fill to none (transparent)
                        if (legend != null && legend.Area != null && legend.Area.FillFormat != null)
                        {
                            legend.Area.FillFormat.FillType = FillType.None;
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
            }
            catch (FileNotFoundException ex)
            {
                Console.Error.WriteLine($"File error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
