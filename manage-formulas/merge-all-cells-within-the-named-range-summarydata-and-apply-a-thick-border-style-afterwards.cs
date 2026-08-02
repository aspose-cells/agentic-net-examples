// Title: Merge a Named Range and Add a Thick Outline Border with Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, define the named range "SummaryData" (B2:D4), merge its cells, and apply a thick black outline border using Aspose.Cells for .NET.
// Keywords: Aspose.Cells merge named range C# | Aspose.Cells outline border | C# merge cells Aspose | UnionRange border Aspose.Cells | Aspose.Cells thick border | named range border C# | Aspose.Cells workbook styling
// Common Searches: Aspose.Cells merge named range C# | How to add border to merged cells Aspose.Cells | Set outline border after merging cells Aspose.Cells | C# Aspose.Cells UnionRange example | Apply thick border to range Aspose.Cells
// Developer Intent: Merge all cells in the named range "SummaryData" and then apply a thick black outline border around the merged area using Aspose.Cells for .NET.
// Use Cases: Create a report title that spans several columns and is highlighted with a thick border for visual emphasis. | Design an invoice header where the company name occupies a merged block outlined to separate it from the body. | Build a dashboard widget that merges cells for a summary label and adds a distinct border to improve readability.
// AI Prompts: Generate C# code with Aspose.Cells that merges a named range and sets a thick red outline border. | Explain how UnionRange can be used to apply outline borders after merging cells in an Aspose.Cells workbook. | Provide a step‑by‑step tutorial for creating, merging, and bordering a named range using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsMergeAndBorderDemo
{
    // Shows how to create a workbook, define the named range "SummaryData" (B2:D4), merge its cells, and apply a thick black outline border using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate data that will be part of the named range
                worksheet.Cells["B2"].PutValue("Data 1");
                worksheet.Cells["C2"].PutValue("Data 2");
                worksheet.Cells["D2"].PutValue("Data 3");
                worksheet.Cells["B3"].PutValue("Data 4");
                worksheet.Cells["C3"].PutValue("Data 5");
                worksheet.Cells["D3"].PutValue("Data 6");
                worksheet.Cells["B4"].PutValue("Data 7");
                worksheet.Cells["C4"].PutValue("Data 8");
                worksheet.Cells["D4"].PutValue("Data 9");

                // Create a range covering B2:D4 and assign it the name "SummaryData"
                AsposeRange summaryRange = worksheet.Cells.CreateRange("B2", "D4");
                summaryRange.Name = "SummaryData";

                // Merge all cells within the named range
                summaryRange.Merge();

                // Convert the merged range to a UnionRange to apply outline borders
                UnionRange unionRange = summaryRange.UnionRanges(new AsposeRange[] { summaryRange });

                // Apply a thick black border around the merged range
                unionRange.SetOutlineBorders(CellBorderType.Thick, Color.Black);

                // Save the workbook
                workbook.Save("MergedSummaryData.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
