// Title: C# Example: Insert a Column Chart into a 3‑Row × 5‑Column Range with Aspose.Cells
// Description: This Aspose.Cells for .NET snippet shows how to create a workbook, add sample data, compute the lower‑right cell indices, and place a column chart that exactly spans three rows and five columns (starting at Excel row 6, column C). The chart is added via Worksheet.Charts.Add(ChartType.Column, topRow, leftColumn, bottomRow, rightColumn), linked to the data range A2:B5, and saved as ChartInRange.xlsx.
// Keywords: Aspose.Cells | C# chart insertion | insert chart specific range | chart size rows columns | Worksheet.Charts.Add | column chart Aspose | Excel chart placement programmatically | Aspose.Cells example GitHub | C# Excel chart bounds | Aspose.Cells chart dimensions
// Common Searches: Aspose.Cells insert chart into specific cells | C# set chart size rows columns Aspose.Cells | place column chart at row 6 column C using Aspose.Cells | define chart bounds topRow leftColumn bottomRow rightColumn Aspose | Aspose.Cells chart placement example
// Developer Intent: Insert a column chart that occupies a defined 3‑row by 5‑column area in an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Add a sales‑summary chart to a pre‑formatted report section without overlapping other data. | Build a dashboard where each chart aligns to a cell grid for a uniform layout across multiple sheets. | Automate chart insertion into a template workbook, guaranteeing a fixed cell block for consistent printing.
// AI Prompts: Generate C# code with Aspose.Cells to insert a line chart into cells D10:G12 and bind it to data range A1:B5. | Show how to resize an existing Aspose.Cells chart so it covers 4 rows and 6 columns starting at row 8, column B. | Write a method that retrieves a chart's current cell range, moves it to a new top‑left position, and keeps its original size.

using Aspose.Cells;
using Aspose.Cells.Charts;

// This Aspose.Cells for .NET snippet shows how to create a workbook, add sample data, compute the lower‑right cell indices, and place a column chart that exactly spans three rows and five columns (starting at Excel row 6, column C). The chart is added via Worksheet.Charts.Add(ChartType.Column, topRow, leftColumn, bottomRow, rightColumn), linked to the data range A2:B5, and saved as ChartInRange.xlsx.
class InsertChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data that the chart will use
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 5; i++)
        {
            sheet.Cells[$"A{i}"].PutValue("Item " + (i - 1));
            sheet.Cells[$"B{i}"].PutValue((i - 1) * 10);
        }

        // Define the upper‑left corner of the chart (zero‑based indices)
        int topRow = 5;      // corresponds to Excel row 6
        int leftColumn = 2;  // corresponds to Excel column C

        // Calculate the lower‑right corner so the chart occupies 3 rows × 5 columns
        int bottomRow = topRow + 2;      // 3 rows total (topRow, topRow+1, topRow+2)
        int rightColumn = leftColumn + 4; // 5 columns total (leftColumn … leftColumn+4)

        // Add a column chart within the specified cell range using the Add(ChartType, int, int, int, int) overload
        int chartIndex = sheet.Charts.Add(ChartType.Column, topRow, leftColumn, bottomRow, rightColumn);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("=Sheet1!$A$2:$B$5", true);

        // Save the workbook to a file
        workbook.Save("ChartInRange.xlsx", SaveFormat.Xlsx);
    }
}
