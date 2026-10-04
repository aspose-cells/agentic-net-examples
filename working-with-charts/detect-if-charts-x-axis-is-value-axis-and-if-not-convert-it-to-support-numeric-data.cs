// Title: Detect and convert an Excel chart's X (Category) axis to a Value axis using Aspose.Cells for .NET
// AI Prompts: Generate C# code that inspects each chart in a workbook and changes the X axis from Category to Value with Aspose.Cells. | Provide a method that iterates worksheets and charts, checks the axis type, and sets it to AxisType.Value, including a fallback for older Aspose.Cells versions. | Show how to load an XLSX file, programmatically switch chart X axes to numeric, and save the updated workbook using Aspose.Cells.
// Common Searches: Aspose.Cells how to change chart X axis to numeric in C# | C# code to convert Excel chart category axis to value axis with Aspose | detect chart axis type and modify it using Aspose.Cells .NET | set chart X axis type programmatically in an existing workbook Aspose.Cells
// Tags: Aspose.Cells set chart X axis to value | C# convert chart category axis to numeric | reflection handle AxisType property Aspose.Cells | iterate worksheets modify chart axes Aspose | Excel chart axis type change .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an XLSX workbook, loops through all worksheets and charts, uses reflection to locate the AxisType (or Type) property of each chart's CategoryAxis, sets it to AxisType.Value, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet and its charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    // Convert the X (Category) axis to a Value axis if possible
                    Axis catAxis = chart.CategoryAxis;
                    if (catAxis != null)
                    {
                        // Use reflection to handle different API versions
                        var axisTypeProp = catAxis.GetType().GetProperty("AxisType");
                        if (axisTypeProp != null && axisTypeProp.CanWrite)
                        {
                            axisTypeProp.SetValue(catAxis, AxisType.Value);
                        }
                        else
                        {
                            var typeProp = catAxis.GetType().GetProperty("Type");
                            if (typeProp != null && typeProp.CanWrite)
                            {
                                typeProp.SetValue(catAxis, AxisType.Value);
                            }
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
