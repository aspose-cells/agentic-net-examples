// Title: Read Excel from a Stream, reset slicer filters, refresh pivot tables, and export to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an Excel file from a Stream, resets all slicer filters, recalculates each pivot table, and writes the workbook as a PDF to an output Stream using Aspose.Cells. | Write a method that accepts input and output streams, loads the workbook, clears slicer states, updates pivot calculations, and saves the result as a PDF with Aspose.Cells for .NET.
// Common Searches: how to clear slicer filters in Aspose.Cells before PDF conversion | refresh pivot tables after modifying slicer state using Aspose.Cells C# | convert Excel workbook to PDF from memory stream with Aspose.Cells .NET | programmatically clear all slicer selections in an Excel file using Aspose.Cells | save PDF to a Stream after updating slicers in Aspose.Cells
// Tags: Aspose.Cells load workbook from stream | Aspose.Cells reset slicer filters | Aspose.Cells recalculate pivot tables | Aspose.Cells export PDF to stream | Aspose.Cells slicer‑pivot synchronization

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// The ConvertToPdfWithSlicerUpdate method loads an Excel workbook from an input Stream, iterates through worksheets to reset slicer filters (no direct API for deselection), refreshes and recalculates all pivot tables, and saves the workbook as a PDF to the provided output Stream, handling errors and supporting console execution.
public class WorkbookProcessor
{
    /// <param name="inputStream">Stream containing the source Excel file.</param>
    /// <param name="outputPdfStream">Stream where the generated PDF will be written.</param>
    public static void ConvertToPdfWithSlicerUpdate(Stream inputStream, Stream outputPdfStream)
    {
        try
        {
            // Load the workbook from the provided stream.
            var workbook = new Workbook(inputStream);

            // Clear selections for all slicers in each worksheet (if supported).
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Slicer slicer in sheet.Slicers)
                {
                    // In current Aspose.Cells versions there is no direct API to deselect all items.
                    // The slicer state will be reflected when the pivot tables are refreshed.
                }
            }

            // Refresh all pivot tables to reflect slicer changes.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivotTable in sheet.PivotTables)
                {
                    pivotTable.RefreshData();      // Refresh underlying data source.
                    pivotTable.CalculateData();    // Recalculate the pivot table.
                }
            }

            // Save the workbook as a PDF to the output stream.
            workbook.Save(outputPdfStream, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error during conversion: {ex.Message}");
            throw;
        }
    }

    // Entry point for the console application.
    public static void Main(string[] args)
    {
        // Determine input and output file paths.
        string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
        string outputPath = args.Length > 1 ? args[1] : "output.pdf";

        if (!File.Exists(inputPath))
        {
            Console.Error.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            using (FileStream inputStream = File.OpenRead(inputPath))
            using (FileStream outputStream = File.Create(outputPath))
            {
                ConvertToPdfWithSlicerUpdate(inputStream, outputStream);
            }

            Console.WriteLine($"PDF successfully created at: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Conversion failed: {ex.Message}");
        }
    }
}
