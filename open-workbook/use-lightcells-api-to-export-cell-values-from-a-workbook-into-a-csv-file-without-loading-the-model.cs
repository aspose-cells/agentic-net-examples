// Title: Export Excel to CSV with Aspose.Cells LightCells ConversionUtility (C#) without loading the workbook model
// Description: Shows how to convert an existing XLSX file to a CSV file using Aspose.Cells.Utility.ConversionUtility, which employs the LightCells engine to export cell values while keeping memory usage minimal.
// Keywords: Aspose.Cells | C# | LightCells | ConversionUtility | XLSX to CSV | export cell values | memory‑efficient conversion | streaming Excel to CSV | large workbook conversion | no workbook model load
// Common Searches: Aspose.Cells convert xlsx to csv without loading workbook | LightCells export Excel to CSV C# | ConversionUtility example for CSV export | memory‑low Excel to CSV conversion Aspose | how to use Aspose.Cells Utility ConversionUtility
// Developer Intent: Convert an existing Excel workbook to a CSV file while avoiding the overhead of loading the full workbook model into memory.
// Use Cases: Generate CSV reports from massive Excel files on a server with limited RAM. | Process user‑uploaded spreadsheets to CSV for downstream analytics without full model parsing. | Automate batch conversions in cloud services where memory consumption must stay low.
// AI Prompts: Provide C# code that uses Aspose.Cells.Utility.ConversionUtility to convert an XLSX file to CSV without opening the workbook model. | Explain how LightCells in Aspose.Cells enables low‑memory Excel‑to‑CSV conversion. | Show a streaming approach for converting large Excel files to CSV with Aspose.Cells LightCells API.

// Example: Export cell values from an existing workbook to a CSV file
// without manually loading the workbook model, using Aspose.Cells ConversionUtility.

// Create a sample workbook and save it as XLSX (this part is just for demonstration;
// in real scenarios the source file already exists).
var sourcePath = "sample.xlsx";
var csvPath = "exported.csv";

var workbook = new Aspose.Cells.Workbook();
var sheet = workbook.Worksheets[0];

// Populate some data
sheet.Cells["A1"].PutValue("Name");
sheet.Cells["B1"].PutValue("Age");
sheet.Cells["A2"].PutValue("John");
sheet.Cells["B2"].PutValue(30);
sheet.Cells["A3"].PutValue("Alice");
sheet.Cells["B3"].PutValue(25);

// Save the workbook (model is loaded only here for the demo)
workbook.Save(sourcePath, Aspose.Cells.SaveFormat.Xlsx);

// Convert the XLSX file to CSV.
// ConversionUtility handles the conversion internally without requiring the caller
// to load the workbook model into memory.
Aspose.Cells.Utility.ConversionUtility.Convert(sourcePath, csvPath);

// At this point 'exported.csv' contains the cell values from the original workbook.
