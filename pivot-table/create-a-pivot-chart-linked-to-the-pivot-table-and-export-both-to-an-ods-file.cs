// Title: Create a Pivot Chart Linked to a Pivot Table and Export to ODS with Aspose.Cells for .NET
// Description: Shows how to build a workbook, add sample data, create a pivot table, generate a column chart linked to that table, configure OdsSaveOptions to keep pivot data, and save the result as an ODS document using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# pivot chart | pivot table ODS export | OdsSaveOptions | pivot chart Aspose | export to ODS | .NET LibreOffice compatibility | chart from pivot table | Aspose.Cells .NET | pivot chart linked to pivot table
// Common Searches: Aspose.Cells create pivot chart C# | Export pivot table and chart to ODS using Aspose | Set PivotSource for chart Aspose.Cells | Save workbook with pivot tables as ODS | C# OdsSaveOptions include pivot tables | Link chart to pivot table in Aspose.Cells
// Developer Intent: Generate a pivot chart tied to a pivot table and save the workbook as an ODS file.
// Use Cases: Produce a financial summary workbook where a pivot table aggregates sales data and a column chart visualizes totals, then export to ODS for LibreOffice sharing. | Automate a reporting pipeline that adds a pivot table and its associated chart to a template workbook and outputs an ODS file for downstream processing. | Create an ODS package containing both pivot data and a linked chart for seamless exchange with open‑source office suites.
// AI Prompts: Write C# code with Aspose.Cells that creates a pivot table from a range, adds a column chart linked to the pivot table, and saves the workbook as an ODS file preserving the pivot structures. | Explain how to configure OdsSaveOptions in Aspose.Cells to ensure pivot tables and pivot charts are retained when exporting to ODS. | Show the steps to set the PivotSource property of a Chart object to reference a specific pivot table in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Ods; // for OdsSaveOptions

// Shows how to build a workbook, add sample data, create a pivot table, generate a column chart linked to that table, configure OdsSaveOptions to keep pivot data, and save the result as an ODS document using Aspose.Cells for C#.
public class PivotChartToOds
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("A");
        worksheet.Cells["B4"].PutValue(30);
        worksheet.Cells["A5"].PutValue("B");
        worksheet.Cells["B5"].PutValue(40);

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("=A1:B5", "D1", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field

        // Add a chart and link it to the pivot table (pivot chart)
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];
        chart.PivotSource = "PivotTable1";   // Set the pivot source for the chart
        chart.RefreshPivotData();            // Refresh chart data from the pivot table

        // Prepare ODS save options (include pivot tables)
        OdsSaveOptions saveOptions = new OdsSaveOptions
        {
            IgnorePivotTables = false // Ensure pivot tables are saved
        };

        // Ensure output directory exists
        string outputPath = "PivotChartDemo.ods";
        string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        // Save the workbook as an ODS file with the specified options
        workbook.Save(outputPath, saveOptions);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
