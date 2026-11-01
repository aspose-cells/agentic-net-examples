// Title: How to embed Excel worksheet images as Base64 in HTML using Aspose.Cells HtmlSaveOptions (C#)
// Common Searches: Aspose.Cells C# export workbook to HTML with images embedded as Base64 | C# HtmlSaveOptions ExportImagesAsBase64 example for self‑contained HTML output | Save Excel file as a single HTML page with embedded pictures using Aspose.Cells | Embedding worksheet pictures in HTML output from Aspose.Cells C# | Generate HTML from Excel without external image files using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | C# export Excel to self-contained HTML | embed Excel images as Base64 in HTML | HTML output with embedded images Aspose.Cells | save workbook as HTML with base64 images C#

// Create a new workbook (or load an existing one)
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(); // using the create rule

// (Optional) Add some data or an image to the workbook here
// For example, insert a picture into the first worksheet
// Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];
// int pictureIndex = sheet.Pictures.Add(5, 5, "sample.png");

// Configure HTML save options to embed images as Base64
Aspose.Cells.HtmlSaveOptions htmlOptions = new Aspose.Cells.HtmlSaveOptions(Aspose.Cells.SaveFormat.Html);
htmlOptions.ExportImagesAsBase64 = true; // Enable embedding images directly in the HTML

// Save the workbook as an HTML file with the specified options
workbook.Save("output.html", htmlOptions); // using the save rule

// The resulting HTML file will contain images encoded as Base64 strings,
 // making it a self‑contained document without external image files.
