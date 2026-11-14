// Title: Apply locale‑specific currency formatting to a smart marker using Aspose.Cells for .NET
// Common Searches: how to format smart marker values as currency with locale in Aspose.Cells C# | Aspose.Cells .NET custom number format for smart markers using en-US locale | C# example of applying [$-en-US] currency format to smart marker data | using WorkbookDesigner to set locale specific currency format in Excel export | smart marker currency formatting with custom number format string in Aspose.Cells
// Tags: smart marker custom number format | locale specific currency formatting Aspose.Cells | WorkbookDesigner apply number format | Excel export currency format C# | Aspose.Cells smart marker locale

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Insert sample data that will be used by the smart marker
// Row 0: Header
sheet.Cells["A1"].PutValue("Product");
sheet.Cells["B1"].PutValue("Amount");

// Row 1..3: Data rows
sheet.Cells["A2"].PutValue("Widget");
sheet.Cells["B2"].PutValue(1234.56);
sheet.Cells["A3"].PutValue("Gadget");
sheet.Cells["B3"].PutValue(7890.12);
sheet.Cells["A4"].PutValue("Doohickey");
sheet.Cells["B4"].PutValue(345.67);

// Add a smart marker row that will apply a locale‑specific currency format
// The smart marker uses a custom number format with a locale identifier.
// Example uses US dollars (en‑US). Change the locale code and currency symbol as needed.
sheet.Cells["A6"].PutValue("&=Product");
sheet.Cells["B6"].PutValue("&=Amount : \"[$$-en-US]#,##0.00\"");

// Create a WorkbookDesigner to process the smart markers
Aspose.Cells.WorkbookDesigner designer = new Aspose.Cells.WorkbookDesigner(workbook);

// Process the smart markers – this will replace them with the actual data
// and apply the custom number format defined above.
designer.Process();

// Save the resulting workbook
workbook.Save("SmartMarkerCurrencyFormatted.xlsx");
