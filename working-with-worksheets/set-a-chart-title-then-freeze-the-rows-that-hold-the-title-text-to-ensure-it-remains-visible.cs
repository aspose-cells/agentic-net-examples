// Title: Set Chart Title and Freeze Header Row in Aspose.Cells (C#)
// Description: Creates a new workbook, populates sample data, inserts a column chart with a visible title, freezes the first worksheet row via PageSetup.PrintTitleRows, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells C# chart title | freeze first row Aspose.Cells | PrintTitleRows example | Excel header freeze Aspose.Cells | column chart Aspose.Cells | repeat header row Aspose.Cells | Aspose.Cells worksheet freeze
// Common Searches: Aspose.Cells set chart title C# | How to freeze first row in Excel using Aspose.Cells | PrintTitleRows C# example | Add column chart with title Aspose.Cells | Repeat header row on printed pages Aspose.Cells
// Developer Intent: Add a chart with a title and keep the worksheet’s header row visible while scrolling or printing.
// Use Cases: Generate a sales report where a column chart displays a title and the header row stays in view during scrolling. | Build an Excel dashboard that repeats the first row on every printed page while showing a titled chart. | Create a distributable workbook with a chart title and frozen header for improved readability.
// AI Prompts: Generate C# code that adds a column chart, sets its title, and freezes the first row using Aspose.Cells. | Explain how PrintTitleRows works for freezing rows and repeating them on printed pages in Aspose.Cells. | Show how to keep a chart title visible while freezing the header row in an Aspose.Cells worksheet (C#).

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, populates sample data, inserts a column chart with a visible title, freezes the first worksheet row via PageSetup.PrintTitleRows, and saves the file as an Excel workbook.
class ChartTitleAndFreezeRows
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Prepare sample data for the chart
        // -------------------------------------------------
        Cells cells = sheet.Cells;
        cells["A1"].PutValue("Product");
        cells["B1"].PutValue("Sales");
        cells["A2"].PutValue("Apple");
        cells["A3"].PutValue("Banana");
        cells["A4"].PutValue("Cherry");
        cells["B2"].PutValue(120);
        cells["B3"].PutValue(85);
        cells["B4"].PutValue(150);

        // -------------------------------------------------
        // Add a chart to the worksheet
        // -------------------------------------------------
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // -------------------------------------------------
        // Set the chart title
        // -------------------------------------------------
        chart.Title.Text = "Quarterly Sales";
        chart.Title.IsVisible = true; // Ensure the title is displayed

        // -------------------------------------------------
        // Freeze the row that contains the worksheet title (A1)
        // -------------------------------------------------
        // The title text is placed in the first row (row 1). 
        // Setting PrintTitleRows repeats this row on every printed page
        // and also freezes it in the UI when scrolling.
        sheet.PageSetup.PrintTitleRows = "$1:$1";

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("ChartWithTitleAndFrozenRow.xlsx");
    }
}
