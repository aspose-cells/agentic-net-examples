// Title: Set Arial for Latin text and MS Mincho for Japanese text in specific Excel cells using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates an Excel workbook, applies the Arial font to cells A1:B1 for Latin characters, and sets the MS Mincho font for cell B1 to display Japanese text, using Aspose.Cells styling APIs. | Show how to define separate Style objects for Western and Far‑East fonts and assign them to ranges or individual cells in a .NET workbook with Aspose.Cells.
// Common Searches: how to apply different fonts for western and CJK characters in Aspose.Cells .NET | Aspose.Cells set Japanese font MS Mincho for a single cell | C# Aspose.Cells change font to Arial for a range of cells | apply Arial font to a range and a different FarEast font to another cell using Aspose.Cells | multilingual font styling in Excel with Aspose.Cells for .NET
// Tags: Aspose.Cells set western font Arial | Aspose.Cells apply FarEast font MS Mincho | C# style cells with different fonts | multilingual font styling in Excel | Excel workbook font assignment using Aspose.Cells

using Aspose.Cells;
using System;

// Demonstrates creating a new workbook, defining separate Style objects for Latin (Arial) and Far‑East (MS Mincho) text, applying the styles to a range and an individual cell, inserting English and Japanese strings, and saving the file as MultilingualFonts.xlsx.
class SetMultilingualFonts
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Latin (Western) style ----------
            Style latinStyle = workbook.CreateStyle();
            latinStyle.Font.Name = "Arial";

            // Apply the Latin style to the range A1:B1
            Aspose.Cells.Range latinRange = sheet.Cells.CreateRange("A1:B1");
            latinRange.ApplyStyle(latinStyle, new StyleFlag() { Font = true });

            // Put Latin text
            sheet.Cells["A1"].PutValue("Hello");

            // ---------- Far East (CJK) style ----------
            Style feStyle = workbook.CreateStyle();
            // Set a CJK font for the style
            feStyle.Font.Name = "MS Mincho";

            // Apply the Far East style to cell B1
            sheet.Cells["B1"].SetStyle(feStyle);

            // Put Japanese text
            sheet.Cells["B1"].PutValue("こんにちは");

            // Save the workbook
            string outputPath = "MultilingualFonts.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
