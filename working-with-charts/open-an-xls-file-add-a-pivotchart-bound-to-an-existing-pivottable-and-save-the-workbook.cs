// Title: Add a PivotChart to an existing XLS workbook and bind it to a PivotTable using Aspose.Cells for .NET
// Common Searches: C# code to insert a PivotChart into an existing XLS file with Aspose.Cells | how to bind a column chart to a pivot table in Aspose.Cells .NET | Aspose.Cells example for creating and positioning a PivotChart in an XLS workbook | load XLS, add PivotChart, save workbook using Aspose.Cells C# | Aspose.Cells .NET add chart to worksheet and link to pivot table
// Tags: pivotchart creation in XLS using Aspose.Cells | column chart pivot source assignment Aspose.Cells | chart positioning rows columns Aspose.Cells | loading and saving XLS workbook Aspose.Cells | setting chart title Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

// The sample loads an existing XLS workbook, checks for a PivotTable on the first worksheet, adds a column PivotChart at a specified range, assigns the PivotTable as the chart's data source, sets a chart title, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.xls";

            // Verify the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one pivot table on the worksheet
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Retrieve the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Define chart position
            int upperLeftRow = 5;
            int upperLeftColumn = 0;
            int lowerRightRow = 20;
            int lowerRightColumn = 10;

            // Add a new column chart and obtain its reference
            int chartIndex = worksheet.Charts.Add(ChartType.Column, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
            Chart chart = worksheet.Charts[chartIndex];

            // Bind the chart to the pivot table (pivot chart)
            chart.PivotSource = pivotTable.Name;

            // Set chart title
            chart.Title.Text = "PivotChart Example";

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
