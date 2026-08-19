// Title: CopyStyle Text‑Wrap Inheritance with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to enable text wrapping on a source range, copy the style to another range using CopyStyle, auto‑fit rows, and save the workbook as an Excel file.
// Keywords: Aspose.Cells CopyStyle | text wrap inheritance | C# Excel style copy | auto fit rows Aspose | wrap long text cells | Aspose.Cells .NET example
// Common Searches: Aspose.Cells copy style with text wrap | how to inherit wrap formatting in Excel using C# | CopyStyle method text wrapping Aspose | auto‑fit rows after copying style Aspose.Cells | C# example for wrapping long text in cells
// Developer Intent: Copy a predefined style that includes text wrapping from one cell range to another programmatically with Aspose.Cells.
// Use Cases: Apply a template‑defined wrap style to dynamically generated report sections. | Ensure description columns in exported Excel sheets automatically wrap without per‑cell styling. | Reuse styled ranges (font, borders, wrap) across multiple worksheets in a single workbook.
// AI Prompts: Write C# code that copies a style with text wrapping from range A1:A3 to range C1:C3 using Aspose.Cells. | Explain the effect of CopyStyle on text‑wrap, borders, and number formats in Aspose.Cells. | Provide a step‑by‑step tutorial for applying wrap inheritance and then auto‑fitting rows in an Aspose.Cells workbook.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsTextWrapInheritance
{
    // Demonstrates how to enable text wrapping on a source range, copy the style to another range using CopyStyle, auto‑fit rows, and save the workbook as an Excel file.
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

                // Populate source cells with long text that needs wrapping
                cells["A1"].PutValue("This is a very long text that should be wrapped within the cell A1.");
                cells["A2"].PutValue("Another long text entry that demonstrates automatic text wrapping in cell A2.");
                cells["A3"].PutValue("Yet another example of a lengthy string that will wrap when the style is applied to cell A3.");

                // Create a source style with text wrapping enabled
                Style sourceStyle = workbook.CreateStyle();
                sourceStyle.IsTextWrapped = true;

                // Apply the source style to a source range (A1:A3)
                AsposeRange sourceRange = cells.CreateRange("A1:A3");
                sourceRange.SetStyle(sourceStyle);

                // Populate destination cells with long text as well
                cells["C1"].PutValue("Destination cell C1 contains a long text that should wrap after copying style.");
                cells["C2"].PutValue("Destination cell C2 also has lengthy content requiring wrapping.");
                cells["C3"].PutValue("Destination cell C3 demonstrates the inheritance of wrap style via CopyStyle.");

                // Define destination range and copy style from source range
                AsposeRange destinationRange = cells.CreateRange("C1:C3");
                destinationRange.CopyStyle(sourceRange);

                // Auto-fit rows to display wrapped text properly
                worksheet.AutoFitRows();

                // Determine output file path
                string outputPath = "TextWrapInheritanceDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
