// Title: Insert an HTML <img> tag into a merged Excel cell using Aspose.Cells for .NET
// AI Prompts: Generate C# code that merges cells A1:C1, sets the cell's HtmlString to an <img> element pointing to a URL or base64 data, and saves the workbook as XLSX using Aspose.Cells. | Provide a step‑by‑step example of inserting an HTML image tag into a merged Excel range with Aspose.Cells for .NET, then persisting the file.
// Common Searches: Aspose.Cells C# merge cells and show image using HTML tag | set HTML content of merged cell in Excel with Aspose.Cells .NET | display picture in merged range via HTML string Aspose.Cells example | C# Aspose.Cells embed external image in merged cells
// Tags: merged cell HtmlString image Aspose.Cells | C# Aspose.Cells insert <img> tag into Excel | save workbook with embedded HTML picture Aspose | base64 image rendering in merged Excel cell Aspose | Excel range merge then HTML image insertion Aspose

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Merge cells A1:C1 (row 0, columns 0 to 2)
sheet.Cells.Merge(0, 0, 1, 3);

// Embed an image using the HTML property of the merged cell
// The <img> tag can reference a local file, a URL, or a base‑64 encoded image.
// Here we use a URL as an example.
sheet.Cells[0, 0].HtmlString = "<img src='https://example.com/sample-image.png' style='width:100%;height:auto;'/>";

// Save the workbook to an XLSX file
workbook.Save("EmbeddedImageAfterMerge.xlsx");
