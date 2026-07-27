// Title: Export Merged Cells to HTML and Validate Layout with Aspose.Cells for .NET
// Description: This C# example creates a workbook, merges range A1:C2, saves it to HTML using HtmlSaveOptions with ValidateMergedAreas enabled, reloads the HTML file, extracts merged areas via GetMergedAreas, and prints their coordinates to confirm that the merged layout is preserved.
// Keywords: Aspose.Cells | C# | HTML export | merged cells | ValidateMergedAreas | round‑trip conversion | load HTML workbook | CellArea | Excel to HTML | preserve layout
// Common Searches: Aspose.Cells keep merged cells when exporting to HTML | HtmlSaveOptions ValidateMergedAreas usage .NET | How to read merged ranges after loading HTML with Aspose.Cells | C# export Excel merged header to HTML | Verify merged cell coordinates after HTML round‑trip
// Developer Intent: Generate an HTML file from a workbook that contains merged cells and ensure the merged structure remains intact after re‑import.
// Use Cases: Publish Excel reports with merged headers on web pages without losing formatting | Perform a round‑trip Excel → HTML → Excel conversion while checking merged regions | Detect layout issues before HTML export by enabling ValidateMergedAreas | Automate validation of merged cell ranges in server‑side document pipelines
// AI Prompts: Write C# code that merges A1:C2, saves the workbook to HTML with ValidateMergedAreas, reloads the HTML, and lists merged areas using Aspose.Cells. | Explain the purpose of HtmlSaveOptions.ValidateMergedAreas and demonstrate how to verify merged cell coordinates after loading an HTML file. | Provide a step‑by‑step tutorial for preserving merged cells during Excel‑to‑HTML conversion and validating them on re‑import with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// This C# example creates a workbook, merges range A1:C2, saves it to HTML using HtmlSaveOptions with ValidateMergedAreas enabled, reloads the HTML file, extracts merged areas via GetMergedAreas, and prints their coordinates to confirm that the merged layout is preserved.
class MergedCellsHtmlDemo
{
    static void Main()
    {
        // -------------------- Create workbook with merged cells --------------------
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells A1:C2 (rows 0-1, columns 0-2)
        worksheet.Cells.Merge(0, 0, 2, 3);
        worksheet.Cells["A1"].PutValue("Merged Header");

        // -------------------- Save workbook to HTML --------------------
        string htmlFile = "merged_cells.html";

        // HtmlSaveOptions inherits from SaveOptions, allowing ValidateMergedAreas
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Validate merged areas before saving (optional but ensures layout correctness)
            ValidateMergedAreas = true
            // MergeEmptyTdType left as default to keep Excel‑like grid lines
        };

        workbook.Save(htmlFile, htmlOptions);

        // -------------------- Load HTML back and verify merged layout --------------------
        Workbook loadedWorkbook = new Workbook(htmlFile);
        Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

        // Retrieve merged areas from the loaded worksheet
        CellArea[] mergedAreas = loadedWorksheet.Cells.GetMergedAreas();

        // Output verification results
        Console.WriteLine($"Number of merged areas after loading HTML: {mergedAreas.Length}");
        foreach (CellArea area in mergedAreas)
        {
            Console.WriteLine($"Merged area: StartRow={area.StartRow}, StartColumn={area.StartColumn}, EndRow={area.EndRow}, EndColumn={area.EndColumn}");
        }
    }
}
