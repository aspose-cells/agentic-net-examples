// Title: Insert a review comment into cell D4 of an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add comment to cell D4 in Excel workbook | how to insert a note in a specific Excel cell using Aspose.Cells .NET | C# example for adding a worksheet comment with Aspose.Cells | Aspose.Cells comment API usage for Excel file generation
// Tags: Aspose.Cells add cell comment C# | Excel worksheet comment insertion Aspose.Cells | C# generate Excel file with cell notes | Aspose.Cells comment API example

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Get the first worksheet (index 0)
Aspose.Cells.Worksheet worksheet = workbook.Worksheets[0];

// Reference cell D4 (row 3, column 3 – zero‑based indexing)
Aspose.Cells.Cell targetCell = worksheet.Cells["D4"];

// Add a comment to the cell
int commentIndex = worksheet.Comments.Add(targetCell.Row, targetCell.Column);
Aspose.Cells.Comment comment = worksheet.Comments[commentIndex];
comment.Note = "Review required before final submission";

// Save the workbook to a file
workbook.Save("CommentExample.xlsx");
