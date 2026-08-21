// Title: Export Excel Workbook with Slicers to PDF and Verify Slicer Geometry using Aspose.Cells for .NET
// Description: Creates a workbook with sample data, adds a table and a slicer, records the slicer's Top/Left/Width/Height, saves the file as XLSX, exports it to PDF, reloads the XLSX, extracts the slicer geometry again, and compares the two sets of coordinates to ensure they match.
// Keywords: Aspose.Cells | C# | .NET | slicer export to PDF | slicer geometry | compare slicer positions | Excel to PDF conversion | preserve slicer layout | PdfSaveOptions | ListObject slicer
// Common Searches: export workbook with slicers to pdf using aspose.cells | get slicer coordinates after saving excel file c# | compare slicer position before and after workbook save | asp.net aspose.cells slicer PDF rendering | verify slicer shape dimensions in exported pdf
// Developer Intent: Generate a PDF from an Excel workbook that contains slicers and confirm that the slicer shapes retain their original dimensions after the workbook is saved and reloaded.
// Use Cases: Produce PDF reports that keep the exact visual placement of slicers. | Automate regression tests to ensure slicer geometry is unchanged across Excel saves. | Validate that PDF rendering of slicers matches the source Excel layout.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook containing slicers to PDF and then compare slicer geometry before and after saving the XLSX. | Provide a method that returns a list of slicer positions (Top, Left, Width, Height) from a worksheet's SlicerCollection. | Explain how the ExportDocumentStructure option affects PDF output of slicers in Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Rendering;
using Aspose.Cells.Tables; // Needed for ListObject

namespace SlicerPdfExportAndCompare
{
    // Creates a workbook with sample data, adds a table and a slicer, records the slicer's Top/Left/Width/Height, saves the file as XLSX, exports it to PDF, reloads the XLSX, extracts the slicer geometry again, and compares the two sets of coordinates to ensure they match.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create workbook with data, table and slicer ----------
                Workbook workbook = new Workbook(); // create new workbook
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Vegetable");
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["A4"].PutValue("Fruit");
                sheet.Cells["B4"].PutValue(150);
                sheet.Cells["A5"].PutValue("Vegetable");
                sheet.Cells["B5"].PutValue(60);

                // Add a table covering the data range
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                // Use DisplayName instead of Name (compatible with all versions)
                table.DisplayName = "DataTable";

                // Add a slicer for the first column of the table
                int slicerIndex = sheet.Slicers.Add(table, table.ListColumns[0], "C2");
                Slicer slicer = sheet.Slicers[slicerIndex];

                // Capture slicer position before saving
                var originalPositions = GetSlicerPositions(sheet.Slicers);

                // ---------- Save original workbook as XLSX ----------
                string xlsxPath = "OriginalWithSlicer.xlsx";
                workbook.Save(xlsxPath); // save workbook

                // ---------- Export workbook to PDF ----------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true
                    // Slicer shapes are rendered by default
                };
                string pdfPath = "WorkbookWithSlicer.pdf";
                workbook.Save(pdfPath, pdfOptions); // save as PDF

                // ---------- Load the saved XLSX to verify slicer positions ----------
                if (File.Exists(xlsxPath))
                {
                    try
                    {
                        Workbook loadedWorkbook = new Workbook(xlsxPath); // load workbook
                        Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                        var loadedPositions = GetSlicerPositions(loadedSheet.Slicers);

                        // ---------- Compare positions ----------
                        Console.WriteLine("Comparing slicer positions before and after save:");
                        for (int i = 0; i < originalPositions.Count; i++)
                        {
                            var orig = originalPositions[i];
                            var load = loadedPositions[i];
                            Console.WriteLine($"Slicer {i + 1}:");
                            Console.WriteLine($"  Original - Top:{orig.Top}, Left:{orig.Left}, Width:{orig.Width}, Height:{orig.Height}");
                            Console.WriteLine($"  Loaded   - Top:{load.Top}, Left:{load.Left}, Width:{load.Width}, Height:{load.Height}");

                            bool same = Math.Abs(orig.Top - load.Top) < 0.01 &&
                                        Math.Abs(orig.Left - load.Left) < 0.01 &&
                                        Math.Abs(orig.Width - load.Width) < 0.01 &&
                                        Math.Abs(orig.Height - load.Height) < 0.01;

                            Console.WriteLine($"  Positions match: {same}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error loading workbook: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Error: File '{xlsxPath}' was not found.");
                }

                // Optional cleanup (uncomment if desired)
                // File.Delete(xlsxPath);
                // File.Delete(pdfPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to extract slicer shape geometry
        private static List<SlicerGeometry> GetSlicerPositions(SlicerCollection slicers)
        {
            var list = new List<SlicerGeometry>();
            foreach (Slicer s in slicers)
            {
                var shape = s.Shape;
                var geom = new SlicerGeometry
                {
                    Top = shape.Top,
                    Left = shape.Left,
                    Width = shape.Width,
                    Height = shape.Height
                };
                list.Add(geom);
            }
            return list;
        }

        // Simple DTO for slicer geometry
        private class SlicerGeometry
        {
            public double Top { get; set; }
            public double Left { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
        }
    }
}
