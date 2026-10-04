// Title: How to fill an Excel worksheet with category and value arrays for chart source data using Aspose.Cells in C#
// AI Prompts: Write C# code that uses Aspose.Cells to create a new workbook, add a header row (Category, Value), populate columns A and B from a string[] and a double[] array, and save the file. | Generate a loop that iterates over parallel category and numeric arrays and inserts each pair into successive rows of the first worksheet with Aspose.Cells, keeping the header intact. | Modify the example to add extra columns for additional chart series (e.g., "Previous Year") while preserving the existing Category and Value columns, then save the workbook.
// Common Searches: Aspose.Cells C# write array data to Excel cells for chart source | populate Excel sheet with category and numeric values using Aspose.Cells | C# Aspose.Cells create data table for column chart | how to save chart source data workbook with Aspose.Cells in .NET
// Tags: Aspose.Cells fill worksheet from arrays | C# create chart source data Excel | Aspose.Cells write header and rows | save Excel workbook with chart data Aspose

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Get the first worksheet
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Populate header row
sheet.Cells["A1"].PutValue("Category");
sheet.Cells["B1"].PutValue("Value");

// Sample source data for the chart
string[] categories = { "Q1", "Q2", "Q3", "Q4" };
double[] values = { 120.5, 150.0, 130.75, 160.25 };

// Fill the data rows
for (int i = 0; i < categories.Length; i++)
{
    // Category column (A)
    sheet.Cells[i + 2, 0].PutValue(categories[i]);   // Row index is i+2 because rows are 1‑based in Excel
    // Value column (B)
    sheet.Cells[i + 2, 1].PutValue(values[i]);
}

// Save the workbook (adjust the path as needed)
workbook.Save("ChartSourceData.xlsx");
