// Title: Inherit text wrapping with CopyStyle for long text cells in Aspose.Cells .NET
// AI Prompts: Generate C# code that inserts lengthy strings into cells, creates a style with IsTextWrapped enabled, and transfers that style to another range using the CopyStyle method. | Modify the sample to detect the used range dynamically and apply the wrap style to all populated cells with a single CopyStyle call. | Extend the example to keep existing cell formats (fonts, colors) while adding text wrapping through style inheritance.
// Common Searches: Aspose.Cells how to copy text wrap style from one range to another in C# | C# Aspose.Cells apply IsTextWrapped to multiple cells using copy style | auto fit row height after copying wrap style with Aspose.Cells .NET | inherit text wrapping for dynamic cell ranges Aspose.Cells example
// Tags: copystyle text wrap Aspose.Cells | wrap style source range .NET | dynamic range style inheritance C# | auto-fit rows after style copy | preserve existing formatting with copystyle

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Shows how to create a workbook, insert long text, define a wrap style, apply it to a source range, copy the style to a target range with CopyStyle, auto‑fit rows, and save the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Insert long text into cells that need wrapping
            cells["A1"].PutValue("This is a very long text that should wrap automatically within the cell when style is inherited.");
            cells["A2"].PutValue("Another long text example that needs wrapping to be visible properly in the worksheet.");

            // Create a source style with text wrapping enabled
            Style wrapStyle = workbook.CreateStyle();
            wrapStyle.IsTextWrapped = true;

            // Apply the wrap style to a source range (B1:B2)
            AsposeRange sourceRange = cells.CreateRange("B1:B2");
            sourceRange.SetStyle(wrapStyle);

            // Copy the style from the source range to the target range (A1:A2)
            AsposeRange targetRange = cells.CreateRange("A1:A2");
            targetRange.CopyStyle(sourceRange);

            // Auto-fit rows so the wrapped text becomes visible
            worksheet.AutoFitRows();

            // Save the workbook
            string outputPath = "WrappedTextWithCopyStyle.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
