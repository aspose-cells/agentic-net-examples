// Title: Automatically AutoFit all columns in a merged Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an existing Excel file, iterates through each worksheet, calls AutoFitColumns(), and saves the workbook with Aspose.Cells. | Update a workbook‑merging routine in C# to invoke AutoFitColumns on every worksheet after the merge using Aspose.Cells.
// Common Searches: C# Aspose.Cells auto fit columns for every sheet after merging workbooks | how to apply AutoFitColumns to all worksheets in an existing Excel file using Aspose.Cells .NET | resize all columns automatically in a merged workbook with Aspose.Cells C#
// Tags: auto fit columns Aspose.Cells | auto fit all worksheets C# | column width adjustment after workbook merge Aspose.Cells | apply AutoFitColumns to merged Excel file .NET | programmatic column resizing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a merged.xlsx workbook (or creates a new one), loops through each worksheet calling AutoFitColumns() to adjust column widths, and saves the result as merged_autofit.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "merged.xlsx";
        const string outputPath = "merged_autofit.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook with a single worksheet as a fallback.
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
                // Optional: add sample data to demonstrate AutoFit.
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample");
                workbook.Worksheets[0].Cells["B1"].PutValue("Data");
            }

            // Apply AutoFit to all columns in each worksheet.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.AutoFitColumns();
            }

            // Save the updated workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
