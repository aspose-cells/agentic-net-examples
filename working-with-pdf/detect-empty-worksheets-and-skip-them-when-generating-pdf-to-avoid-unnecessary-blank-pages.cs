// Title: Skip empty worksheets when converting an Excel workbook to PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# code to ignore blank sheets during Aspose.Cells Excel to PDF conversion | remove empty worksheets before saving PDF with Aspose.Cells .NET | how to prevent blank pages in PDF generated from Excel using Aspose.Cells | detect worksheets with no data using MaxDataRow MaxDataColumn Aspose.Cells | Aspose.Cells PDF export skip sheets without content C#
// Tags: skip empty worksheets Aspose.Cells PDF conversion | detect blank sheet MaxDataRow MaxDataColumn Aspose.Cells | delete empty worksheets prior to PDF generation .NET | Aspose.Cells workbook.Save PDF without blank pages | filter worksheets by data presence C# Aspose.Cells

// Load the workbook from an existing Excel file
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Iterate through all worksheets and collect indexes of empty sheets
System.Collections.Generic.List<int> emptySheetIndexes = new System.Collections.Generic.List<int>();
for (int i = 0; i < workbook.Worksheets.Count; i++)
{
    Aspose.Cells.Worksheet sheet = workbook.Worksheets[i];

    // A worksheet is considered empty when it has no used cells
    // MaxDataRow and MaxDataColumn return -1 when there is no data
    if (sheet.Cells.MaxDataRow < 0 || sheet.Cells.MaxDataColumn < 0)
    {
        emptySheetIndexes.Add(i);
    }
}

// Remove empty worksheets in reverse order to keep indexes valid
for (int i = emptySheetIndexes.Count - 1; i >= 0; i--)
{
    workbook.Worksheets.RemoveAt(emptySheetIndexes[i]);
}

// Save the workbook as PDF; only non‑empty worksheets will be rendered
workbook.Save("output.pdf", Aspose.Cells.SaveFormat.Pdf);
