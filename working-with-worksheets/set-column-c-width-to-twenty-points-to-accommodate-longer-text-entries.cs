// Title: How to set column C width to 20 points in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, sets column C (zero‑based index 2) width to 20 points with Aspose.Cells, and saves the file. | Create a reusable C# method that receives an Aspose.Cells Worksheet object and adjusts column C width to 20 character units. | Show how to change the width of a specific column by its index in an Aspose.Cells worksheet without using column letters.
// Common Searches: Aspose.Cells C# set column width to 20 characters | How to change width of column C in an Excel file using Aspose.Cells .NET | Set column width by index with Aspose.Cells example C# | Adjust Excel column width to a specific point size using Aspose.Cells for .NET | C# Aspose.Cells column width 20 points worksheet
// Tags: set column width Aspose.Cells | Aspose.Cells column width by index | C# adjust Excel column width | Aspose.Cells worksheet column sizing | column C width 20 points | Aspose.Cells SetColumnWidth example

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet (index 0)
Aspose.Cells.Worksheet worksheet = workbook.Worksheets[0];

// Set the width of column C (zero‑based index 2) to 20 points (approx. 20 character units)
worksheet.Cells.SetColumnWidth(2, 20);

// Save the workbook to a file
workbook.Save("Output.xlsx");
