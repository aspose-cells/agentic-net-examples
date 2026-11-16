// Title: Apply a logarithmic X‑axis to a scatter chart using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set X axis to logarithmic scale for scatter chart | How to enable log base on chart axis with Aspose.Cells .NET | Create scatter chart with logarithmic X values in Excel using Aspose.Cells | Aspose.Cells scatter plot log scale example code C#
// Tags: scatter chart X axis log scaling Aspose.Cells | CategoryAxis.IsLogarithmic property C# | chart axis logarithmic base setting Aspose.Cells | export Excel file with scatter chart log scale | Aspose.Cells data visualization logarithmic axis

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Get the first worksheet
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Fill sample data for the scatter chart (X in column A, Y in column B)
for (int i = 0; i < 10; i++)
{
    // X‑axis values (linear)
    sheet.Cells[i, 0].PutValue(i + 1);
    // Y‑axis values (exponential to illustrate distribution)
    sheet.Cells[i, 1].PutValue(System.Math.Pow(10, i));
}

// Add a scatter chart to the worksheet
int chartIndex = sheet.Charts.Add(Aspose.Cells.Charts.ChartType.Scatter, 5, 0, 20, 10);
Aspose.Cells.Charts.Chart chart = sheet.Charts[chartIndex];

// Add the series: Y‑values range, and then assign X‑values range
chart.NSeries.Add("B1:B10", true);               // Y values
chart.NSeries[0].XValues = "A1:A10";            // X values

// Apply logarithmic scale to the X‑axis (CategoryAxis for scatter charts)
chart.CategoryAxis.IsLogarithmic = true;        // Enable logarithmic scaling
chart.CategoryAxis.LogBase = 10;                // Optional: set base of the log (default is 10)

// Save the workbook with the chart
workbook.Save("ScatterLogScale.xlsx");
