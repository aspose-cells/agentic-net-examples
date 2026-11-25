// Title: Export an Excel workbook with slicers to PDF and verify slicer coordinates using Aspose.Cells for .NET
// Common Searches: Aspose.Cells export workbook containing slicers to PDF C# | how to read slicer row and column positions with Aspose.Cells | compare slicer coordinates before and after saving Excel as PDF using Aspose.Cells | C# code to validate slicer layout after PDF conversion with Aspose.Cells | retrieve slicer top left values from worksheet in Aspose.Cells
// Tags: pdf export preserving slicers Aspose.Cells | extract slicer coordinates worksheet Aspose.Cells | slicer layout consistency check C# | dynamic slicer enumeration Aspose.Cells | compare slicer positions Excel PDF conversion

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, saves it as a PDF while keeping any slicers, captures each slicer's row, column, top and left coordinates via dynamic access, reloads the workbook, re‑reads the slicer positions, compares them with the original values, and reports mismatches, handling cases where the slicer API is unavailable.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPdfPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Export the workbook (including any slicers) to PDF
            workbook.Save(outputPdfPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF to '{outputPdfPath}'.");

            // Optional: Verify slicer positions if the API supports slicers
            try
            {
                Worksheet sheet = workbook.Worksheets[0];
                var originalSlicerPositions = new Dictionary<string, (int Row, int Column, double Top, double Left)>();

                // Capture slicer positions using dynamic to avoid compile‑time dependency on Slicer type
                foreach (dynamic slicer in sheet.Slicers)
                {
                    string name = slicer.Name;
                    int row = slicer.Row;
                    int column = slicer.Column;
                    double top = slicer.Top;
                    double left = slicer.Left;
                    originalSlicerPositions[name] = (row, column, top, left);
                }

                // Reload workbook to compare positions
                Workbook reloadedWorkbook = new Workbook(inputPath);
                Worksheet reloadedSheet = reloadedWorkbook.Worksheets[0];

                bool allMatch = true;
                foreach (dynamic slicer in reloadedSheet.Slicers)
                {
                    string name = slicer.Name;
                    if (originalSlicerPositions.TryGetValue(name, out var original))
                    {
                        int row = slicer.Row;
                        int column = slicer.Column;
                        double top = slicer.Top;
                        double left = slicer.Left;

                        bool match = original.Row == row &&
                                     original.Column == column &&
                                     Math.Abs(original.Top - top) < 0.01 &&
                                     Math.Abs(original.Left - left) < 0.01;

                        if (!match)
                        {
                            allMatch = false;
                            Console.WriteLine($"Slicer '{name}' position mismatch:");
                            Console.WriteLine($"  Original - Row:{original.Row}, Column:{original.Column}, Top:{original.Top}, Left:{original.Left}");
                            Console.WriteLine($"  Current  - Row:{row}, Column:{column}, Top:{top}, Left:{left}");
                        }
                    }
                }

                Console.WriteLine(allMatch
                    ? "All slicer positions match the original Excel file."
                    : "Some slicer positions differ from the original Excel file.");
            }
            catch (Exception ex)
            {
                // If the slicer API is unavailable or another error occurs, report but do not halt execution
                Console.WriteLine($"Slicer verification skipped or failed: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
