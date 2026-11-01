// Title: Insert an HTML <img> tag into a merged Excel cell using Aspose.Cells for .NET (C#)
// Common Searches: how to display an image inside a merged cell with Aspose.Cells C# | using HtmlString property to embed picture in merged Excel cells Aspose.Cells | set HTML content with <img> tag for a merged range in Aspose.Cells .NET | adjust row height and column width for image rendering in merged cells Aspose.Cells | save workbook with merged cell containing an embedded image using Aspose.Cells C#
// Tags: Aspose.Cells HtmlString merged cell image | C# merge range embed HTML Excel | render img tag in Excel Aspose.Cells | configure cell dimensions for HTML image Aspose.Cells | save workbook with merged cell picture Aspose.Cells

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Define the range to merge (e.g., B2:D4)
int firstRow = 1;   // zero‑based index (row 2)
int firstColumn = 1; // column B
int totalRows = 3;   // rows 2‑4
int totalColumns = 3; // columns B‑D

// Merge the cells
sheet.Cells.Merge(firstRow, firstColumn, totalRows, totalColumns);

// Set the HTML string that contains an <img> tag.
// The image can be referenced by a URL or a Base64 data URI.
string imageUrl = "https://example.com/sample-image.png";
string html = $"<img src=\"{imageUrl}\" style=\"width:100%;height:auto;\"/>";

// Assign the HTML string to the top‑left cell of the merged area.
// Aspose.Cells will render the image inside the merged cell.
sheet.Cells[firstRow, firstColumn].HtmlString = html;

// Optionally adjust the row height and column width to better display the picture
sheet.Cells.SetRowHeight(firstRow, 120); // increase height for the merged rows
for (int col = firstColumn; col < firstColumn + totalColumns; col++)
{
    sheet.Cells.SetColumnWidth(col, 30); // widen columns B‑D
}

// Save the workbook
workbook.Save("MergedCellWithImage.xlsx");
