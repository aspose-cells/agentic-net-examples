// Title: How to enable text wrapping for a specific cell range in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a style with IsTextWrapped = true and applies it to the range A1:C3 using Aspose.Cells. | Show how to use Aspose.Cells StyleFlag to apply a text‑wrap style to an entire cell range and then save the workbook. | Explain the steps to define a cell range, enable text wrapping, and export the file as WrappedText.xlsx with Aspose.Cells for .NET.
// Common Searches: asp.net aspose.cells set wrap text for range A1:C3 c# | c# aspose.cells apply text wrapping to multiple cells | how to use StyleFlag to enable text wrap in an Aspose.Cells workbook | save Excel file with wrapped text using Aspose.Cells .NET
// Tags: apply text wrap style Aspose.Cells | Aspose.Cells StyleFlag text wrapping | C# set IsTextWrapped property | wrap text for cell range Excel Aspose | save workbook with wrapped cells Aspose.Cells

// Create a new workbook
var workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
var worksheet = workbook.Worksheets[0];

// Define the cell range to apply text wrap (e.g., A1:C3)
var range = worksheet.Cells.CreateRange("A1:C3");

// Create a style object and enable text wrapping
var style = workbook.CreateStyle();
style.IsTextWrapped = true;

// Apply the style to the defined range
range.ApplyStyle(style, new Aspose.Cells.StyleFlag() { All = true });

// Save the workbook to a file
workbook.Save("WrappedText.xlsx");
