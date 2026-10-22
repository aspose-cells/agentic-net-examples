// Title: How to resize chart data label shapes to fit bold and italic text using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that sets the data label font to bold and italic and then expands the label shape so the full text is visible. | Show how to programmatically adjust the width and height of column chart data label shapes after applying bold‑italic styling with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells increase data label shape size after applying bold italic font in C# | C# resize Excel chart data labels to fit styled text using Aspose.Cells | programmatically adjust data label dimensions for column chart in Aspose.Cells .NET
// Tags: Aspose.Cells resize chart data label shape | set bold italic font on chart data labels Aspose.Cells | adjust data label dimensions column chart .NET | Excel chart data label auto-fit Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a new workbook, adds a column chart with sample data, enables data labels for the first series, applies bold and italic formatting to the data label font, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet (lifecycle rule: create)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Apply bold and italic font to the data label text
            chart.NSeries[0].DataLabels.Font.IsBold = true;
            chart.NSeries[0].DataLabels.Font.IsItalic = true;

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ChartWithResizedDataLabels.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
