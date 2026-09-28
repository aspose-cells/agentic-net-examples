// Title: Set the worksheet right margin to 10 points with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells PageSetup.RightMargin to set a 10‑point right margin for an Excel worksheet in C#. | Create a workbook, configure the right page margin to 10 points, and save the file using Aspose.Cells. | Adjust the right margin of the first worksheet to provide extra space for long sentences with the Aspose.Cells API.
// Common Searches: Aspose.Cells C# set worksheet right margin to 10 points | how to increase right page margin in Excel using Aspose.Cells .NET | C# code to set right margin in points with Aspose.Cells PageSetup | provide extra space for long sentences by adjusting right margin in Aspose.Cells | Aspose.Cells PageSetup.RightMargin property example
// Tags: Aspose.Cells set right margin points | worksheet page setup right margin C# | Excel right margin 10 points Aspose | adjust worksheet margins using Aspose.Cells API | right margin configuration Aspose.Cells .NET

// Create a new workbook
var workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
var worksheet = workbook.Worksheets[0];

// Set the right margin to 10 points (provides space for long sentences)
worksheet.PageSetup.RightMargin = 10f; // RightMargin is measured in points

// Save the workbook to a file
workbook.Save("Output.xlsx");
