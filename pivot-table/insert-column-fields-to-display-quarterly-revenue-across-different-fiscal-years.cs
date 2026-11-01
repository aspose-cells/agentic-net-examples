// Title: Add fiscal year and quarterly revenue columns to an Excel worksheet with Aspose.Cells for .NET
// Common Searches: how to create Excel columns for fiscal year and quarters using Aspose.Cells C# | populate multi-year quarterly revenue data in a worksheet with Aspose.Cells .NET | auto fit Excel columns after writing data with Aspose.Cells in C# | save a workbook as .xlsx after adding custom headers with Aspose.Cells | write a two‑dimensional array to Excel cells using Aspose.Cells C#
// Tags: Aspose.Cells add custom column headers | C# write two‑dimensional array to Excel cells | Aspose.Cells auto‑fit columns .NET | Aspose.Cells save workbook as xlsx | populate fiscal year quarterly revenue Aspose.Cells

// Create a new workbook (using the provided create rule)
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Add a worksheet to hold the revenue data
Aspose.Cells.Worksheet sheet = workbook.Worksheets[workbook.Worksheets.Add()];
sheet.Name = "QuarterlyRevenue";

// Define column headers: Fiscal Year and the four quarters
string[] headers = { "Fiscal Year", "Q1", "Q2", "Q3", "Q4" };
for (int col = 0; col < headers.Length; col++)
{
    // Insert header text into the first row
    sheet.Cells[0, col].PutValue(headers[col]);
    // Apply a simple style for readability (bold)
    Aspose.Cells.Style style = sheet.Cells[0, col].GetStyle();
    style.Font.IsBold = true;
    sheet.Cells[0, col].SetStyle(style);
}

// Sample data: revenue for three fiscal years
// Each inner array corresponds to: Fiscal Year, Q1, Q2, Q3, Q4
object[,] revenueData = new object[,]
{
    { "FY 2022", 125000, 138000, 142500, 150000 },
    { "FY 2023", 130000, 145000, 148000, 155000 },
    { "FY 2024", 135000, 150000, 155000, 160000 }
};

// Populate the worksheet with the sample data starting from row 1
for (int row = 0; row < revenueData.GetLength(0); row++)
{
    for (int col = 0; col < revenueData.GetLength(1); col++)
    {
        sheet.Cells[row + 1, col].PutValue(revenueData[row, col]);
    }
}

// Auto-fit columns for better appearance
sheet.AutoFitColumns();

// Save the workbook (using the provided save rule)
workbook.Save("QuarterlyRevenue.xlsx");
