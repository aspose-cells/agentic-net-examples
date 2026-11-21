// Title: Copy only the formatting of a source range to a destination range with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy range formatting without changing cell values | How to transfer cell styles between two ranges in an Excel workbook using Aspose.Cells | Copy only styles from A1:C3 to E1:G3 with Aspose.Cells for .NET | Preserve data while applying source range formatting to another range in C# Excel library | Use CopyStyle method to duplicate formatting between ranges in Aspose.Cells
// Tags: Aspose.Cells CopyStyle API | Excel range style duplication .NET | preserve data while copying styles Aspose | duplicate formatting from source to target range C# | apply source range styles to another range Aspose.Cells

// Load an existing workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Get the first worksheet (you can change the index or name as needed)
Aspose.Cells.Worksheet worksheet = workbook.Worksheets[0];

// Define the source range whose formatting will be copied
Aspose.Cells.Range sourceRange = worksheet.Cells.CreateRange("A1:C3");

// Define the destination range that will receive the formatting
Aspose.Cells.Range destinationRange = worksheet.Cells.CreateRange("E1:G3");

// Copy only the formatting (styles) from the source range to the destination range
destinationRange.CopyStyle(sourceRange);

// Save the workbook with the applied formatting changes
workbook.Save("output.xlsx");
