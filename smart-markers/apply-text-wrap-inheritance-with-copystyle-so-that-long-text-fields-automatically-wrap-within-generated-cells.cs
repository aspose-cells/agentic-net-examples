// Title: How to inherit text wrap style using CopyStyle for long text cells in Aspose.Cells for .NET
// AI Prompts: Create a C# workbook with Aspose.Cells, define a style with IsTextWrapped = true, and apply it to several cells using SetStyle/GetStyle. | Show how to copy a wrapped‑text style from a source cell to other cells with CopyStyle in Aspose.Cells for .NET. | Write code that inserts long strings into multiple cells, reuses the source cell's wrap style, and saves the workbook as an Excel file.
// Common Searches: aspnet copy cell style with text wrap aspose.cells | inherit IsTextWrapped property when duplicating cell style c# | apply text wrapping to multiple cells using SetStyle in Aspose.Cells | example of using CopyStyle to propagate wrap setting in Excel workbook .NET
// Tags: aspocells style inheritance text wrap | c# aspocells inherit cell style properties | excel .net apply text wrapping to cells | long text cell formatting aspocells | smart markers style copying aspocells

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates creating a workbook, defining a style with IsTextWrapped = true, applying it to a source cell, copying the style to other cells via SetStyle/GetStyle (or CopyStyle), inserting long strings, and saving the file as WrappedTextExample.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Get the first worksheet
                var sheet = workbook.Worksheets[0];

                // Create a style with text wrapping enabled
                var wrapStyle = workbook.CreateStyle();
                wrapStyle.IsTextWrapped = true;

                // Write a long text into cell A1
                var sourceCell = sheet.Cells["A1"];
                sourceCell.PutValue(
                    "This is a very long text that should automatically wrap within the generated cell. " +
                    "It demonstrates how text wrapping can be inherited using CopyStyle in Aspose.Cells.");

                // Apply the wrap style to the source cell
                sourceCell.SetStyle(wrapStyle);

                // Copy the style (including text wrap) to other cells using SetStyle with the source cell's style
                sheet.Cells["B1"].SetStyle(sourceCell.GetStyle());
                sheet.Cells["C1"].SetStyle(sourceCell.GetStyle());

                // Optionally put different long texts into the copied cells
                sheet.Cells["B1"].PutValue(
                    "Another long piece of text placed in B1 to verify that wrapping works correctly after copying the style.");
                sheet.Cells["C1"].PutValue(
                    "Yet another example of a lengthy string in C1, ensuring that the wrap setting is inherited.");

                // Save the workbook
                workbook.Save("WrappedTextExample.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
