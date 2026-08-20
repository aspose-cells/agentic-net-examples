// Title: Merge P2:P5, Wrap Text, Apply Auto‑Filter and Save XLSX using Aspose.Cells for .NET
// Description: Creates a new workbook, adds sample data, sets an auto‑filter on A1:B5, merges cells P2:P5, enables text wrapping for the merged range, and saves the file as MergedWrapAutoFilter.xlsx (XLSX) with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | merge cells | wrap text | auto filter | save as XLSX | Excel export .NET | IsTextWrapped style | worksheet AutoFilter | merged cell formatting
// Common Searches: Aspose.Cells merge cells and wrap text C# | How to apply auto filter with Aspose.Cells .NET | Set IsTextWrapped on merged cells Aspose.Cells | Save workbook as XLSX using Aspose.Cells | C# code for merging P2:P5 and enabling wrap text
// Developer Intent: Generate an XLSX file that contains a filterable table and a merged, wrap‑enabled cell block using Aspose.Cells for .NET.
// Use Cases: Create printable reports with a multi‑row header that needs text wrapping. | Export filtered inventory data with a merged notes column spanning several rows. | Automate sales dashboards where comments span rows and wrap automatically.
// AI Prompts: Write C# code with Aspose.Cells to merge cells P2:P5, enable text wrap, apply an auto‑filter to A1:B5, and save as XLSX. | Show how to set the IsTextWrapped style on a merged range in Aspose.Cells for .NET. | Provide an example that combines merging, wrapping, and auto‑filtering in a single Aspose.Cells workbook.

using System;
using Aspose.Cells;

// Creates a new workbook, adds sample data, sets an auto‑filter on A1:B5, merges cells P2:P5, enables text wrapping for the merged range, and saves the file as MergedWrapAutoFilter.xlsx (XLSX) with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data for auto‑filter demonstration
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("Apple");
        sheet.Cells["B4"].PutValue(15);
        sheet.Cells["A5"].PutValue("Banana");
        sheet.Cells["B5"].PutValue(25);

        // Apply auto filter to the data range
        sheet.AutoFilter.Range = "A1:B5";

        // Merge cells P2:P5 (column P = index 15, rows 2‑5 = indices 1‑4)
        sheet.Cells.Merge(1, 15, 4, 1);

        // Set wrap‑text property for the merged cell
        Style style = sheet.Cells["P2"].GetStyle();
        style.IsTextWrapped = true;
        sheet.Cells["P2"].SetStyle(style);

        // Optional: place long text into the merged cell to see wrapping
        sheet.Cells["P2"].PutValue("This is a long text that should wrap inside the merged cells P2:P5.");

        // Save the workbook as XLSX
        workbook.Save("MergedWrapAutoFilter.xlsx", SaveFormat.Xlsx);
    }
}
