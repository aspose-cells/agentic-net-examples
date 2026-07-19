// Title: Set Custom Currency Format for Chart Series Values with Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, inserts a column chart, and uses the ValuesFormatCode property to format the series values as currency with the pattern "$#,##0". Saves the file as SeriesValuesCustomCurrency.xlsx.
// Keywords: Aspose.Cells | ValuesFormatCode | custom currency format | chart series formatting | C# Excel chart | number format pattern | financial chart export
// Common Searches: Aspose.Cells set chart series number format | How to format chart data as currency in C# | ValuesFormatCode example Aspose.Cells | Custom number format for Excel chart series .NET | Apply currency pattern to chart values using Aspose.Cells
// Developer Intent: Apply a custom currency number format to a chart series' values.
// Use Cases: Display monetary values in financial reports generated with Aspose.Cells. | Standardize currency appearance across multiple Excel charts for accounting dashboards. | Localize chart number formats for different regional settings. | Show whole-dollar amounts in presentation‑ready Excel files.
// AI Prompts: Write C# code using Aspose.Cells to set the ValuesFormatCode of a chart series to "€#,##0.00". | Show how to assign different number formats to each series in a multi‑series chart with Aspose.Cells. | Explain how workbook locale influences custom format strings applied via ValuesFormatCode.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, and uses the ValuesFormatCode property to format the series values as currency with the pattern "$#,##0". Saves the file as SeriesValuesCustomCurrency.xlsx.
class SetSeriesValuesFormatCode
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(1000);
        worksheet.Cells["B3"].PutValue(2000);
        worksheet.Cells["B4"].PutValue(3000);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the series and categories
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Apply a custom currency format to the series values
        chart.NSeries[0].ValuesFormatCode = "$#,##0";

        // Save the workbook to a file
        workbook.Save("SeriesValuesCustomCurrency.xlsx");
    }
}
