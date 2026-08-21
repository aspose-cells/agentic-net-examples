// Title: C# – Set X‑Axis Title on a Column Chart with Aspose.Cells for .NET
// Description: Demonstrates how to create an Excel workbook, add sales data, insert a column chart, define its data range, and assign a visible, descriptive X‑axis (CategoryAxis) title using Aspose.Cells for .NET before saving the file.
// Keywords: Aspose.Cells | .NET | C# | chart axis title | CategoryAxis | X axis label | column chart | Excel workbook | set chart title | sample code | GitHub example
// Common Searches: Aspose.Cells set X axis title C# | how to add category axis label in Aspose.Cells | C# column chart axis title Aspose.Cells | make X‑axis title visible Aspose.Cells .NET | sample code for chart axis title Aspose
// Developer Intent: Add a custom, visible label to the X‑axis of an Excel chart.
// Use Cases: Generate a sales performance report where the X‑axis shows the month range for clarity. | Create financial dashboards that require explicit axis titles to indicate data periods. | Update existing workbooks programmatically to reflect dynamic date ranges on chart axes.
// AI Prompts: Provide C# code that sets CategoryAxis.Title.Text and makes it visible for a line chart using Aspose.Cells. | Show how to set both the chart title and X‑axis title, then hide the X‑axis title on demand. | Explain how to change the X‑axis title based on a variable date range in an Aspose.Cells workbook.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create an Excel workbook, add sales data, insert a column chart, define its data range, and assign a visible, descriptive X‑axis (CategoryAxis) title using Aspose.Cells for .NET before saving the file.
class SetXAxisTitle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["B2"].PutValue(1200);
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["B3"].PutValue(1500);
        sheet.Cells["A4"].PutValue("Mar");
        sheet.Cells["B4"].PutValue(1800);

        // Insert a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories (X‑axis)

        // Set the X‑axis (CategoryAxis) title to a descriptive label
        chart.CategoryAxis.Title.Text = "Months (Jan‑Mar)";
        chart.CategoryAxis.Title.IsVisible = true;

        // (Optional) Set a chart title for completeness
        chart.Title.Text = "Quarterly Sales";
        chart.Title.IsVisible = true;

        // Save the workbook to a file
        workbook.Save("SetXAxisTitle.xlsx");
    }
}
