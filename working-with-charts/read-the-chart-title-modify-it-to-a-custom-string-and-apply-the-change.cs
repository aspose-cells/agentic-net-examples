// Title: Modify an Excel Chart Title using Aspose.Cells for .NET (C#)
// Description: This C# example shows how to create a workbook, add a column chart, read its current title, replace it with a custom string, make the title visible, and save the file as ModifiedChartTitle.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart title C# | change Excel chart title programmatically | read chart title Aspose.Cells | set chart title visibility Aspose.Cells | modify chart title .NET | Aspose.Cells column chart example | update chart title C#
// Common Searches: Aspose.Cells how to change chart title C# | read chart title Aspose.Cells .NET | set custom title for Excel chart using Aspose.Cells | make chart title visible Aspose.Cells | C# update Excel chart title programmatically
// Developer Intent: Read the existing chart title, replace it with custom text, ensure the title is visible, and save the workbook.
// Use Cases: Generate dynamic chart titles based on user input or data context. | Log or display current chart titles before performing batch updates. | Standardize chart titles across multiple worksheets for consistent reporting. | Guarantee chart titles remain visible after automated modifications.
// AI Prompts: Create C# code with Aspose.Cells that reads a chart's title, changes it to a user‑provided string, and saves the workbook. | Show how to check a chart title's visibility, enable it if hidden, and then update the title text using Aspose.Cells for .NET. | Provide a snippet that loops through all charts in a worksheet, prefixes each title with "Q1 -", and saves the changes.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example shows how to create a workbook, add a column chart, read its current title, replace it with a custom string, make the title visible, and save the file as ModifiedChartTitle.xlsx using Aspose.Cells for .NET.
class ModifyChartTitle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set an initial title
        chart.Title.Text = "Original Chart Title";
        chart.Title.IsVisible = true;

        // Read the current title
        string currentTitle = chart.Title.Text;
        Console.WriteLine("Current Title: " + currentTitle);

        // Modify the title to a custom string
        string customTitle = "Custom Chart Title";
        chart.Title.Text = customTitle;

        // Verify the change
        Console.WriteLine("Updated Title: " + chart.Title.Text);

        // Save the workbook
        workbook.Save("ModifiedChartTitle.xlsx");
    }
}
