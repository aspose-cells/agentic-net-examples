// Title: Automatically adjust all worksheet column widths to fit their content using Aspose.Cells AutoFitColumns in C#
// AI Prompts: Generate C# code that loads an existing Excel file with Aspose.Cells, calls sheet.AutoFitColumns(), and saves the workbook. | Create a reusable C# method that accepts a Worksheet object and applies AutoFitColumns to either the whole sheet or a specified column range. | Show how to combine data population and column auto‑fit in a single Aspose.Cells workflow, then export the result to a .xlsx file.
// Common Searches: Aspose.Cells C# AutoFitColumns example for adjusting column width | how to auto resize Excel columns with Aspose.Cells .NET library | programmatically fit column width to cell content using Aspose.Cells | C# code to auto‑fit all columns in a worksheet after writing data | Aspose.Cells AutoFitColumns on loaded workbook and save
// Tags: auto-fit columns Aspose.Cells | adjust worksheet column width C# | Aspose.Cells column auto-fit method | fit Excel column width programmatically .NET | auto resize worksheet columns after data entry

// Create a new workbook (or load an existing one)
// If you need to load an existing file, replace the constructor argument with the file path.
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Populate some sample data (optional – replace with your own data)
sheet.Cells["A1"].PutValue("ID");
sheet.Cells["B1"].PutValue("Name");
sheet.Cells["C1"].PutValue("Description");

sheet.Cells["A2"].PutValue(1);
sheet.Cells["B2"].PutValue("Alice");
sheet.Cells["C2"].PutValue("Senior Engineer with 10 years of experience");

sheet.Cells["A3"].PutValue(2);
sheet.Cells["B3"].PutValue("Bob");
sheet.Cells["C3"].PutValue("Junior Analyst");

// Automatically adjust the width of all columns in the worksheet to fit the content
sheet.AutoFitColumns();

// Save the workbook to a file (adjust the path as needed)
workbook.Save("AdjustedColumns.xlsx");
