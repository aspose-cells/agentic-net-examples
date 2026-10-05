// Title: Auto-fit a column after inserting wrapped multiline text in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that places a newline‑separated string into a cell, enables text wrapping, and then calls AutoFitColumn to resize the column using Aspose.Cells. | Demonstrate how to apply text wrapping to a cell and automatically adjust its column width so all lines are visible in an Excel workbook with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# auto fit column after wrapping multiline text | how to set text wrap and auto‑size column in Aspose.Cells .NET | C# example inserting newline characters in a cell and auto‑adjusting column width with Aspose.Cells | auto fit column width for wrapped cell content using Aspose.Cells for .NET | Aspose.Cells AutoFitColumn not working with wrapped text C#
// Tags: auto‑fit column after text wrap Aspose.Cells | insert multiline text cell C# | enable cell text wrapping Aspose.Cells | adjust column width for wrapped content .NET | Aspose.Cells AutoFitColumn usage | multiline cell handling Aspose.Cells

// Create a new workbook
var workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
var worksheet = workbook.Worksheets[0];

// Insert multiline text into cell A1 (row 0, column 0)
var cell = worksheet.Cells[0, 0];
cell.PutValue("Line 1\nLine 2\nLine 3");

// Enable text wrapping for the cell so the lines are displayed on separate rows
var style = cell.GetStyle();
style.IsTextWrapped = true;
cell.SetStyle(style);

// Auto‑fit the column to make all lines fully visible
worksheet.AutoFitColumn(0);

// Save the workbook
workbook.Save("MultilineAutoFit.xlsx");
