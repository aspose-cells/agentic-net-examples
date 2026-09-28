// Title: Define worksheet print area to match a slicer’s bounds and export the slicer as a PNG image with Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads a slicer's UpperLeftRow/Column and LowerRightRow/Column, sets the worksheet PageSetup.PrintArea to that range, and saves the slicer as a PNG using ImageOrPrintOptions.OnlyArea. | Show how to create a Range object covering a slicer shape and convert it to an image byte array with Aspose.Cells, then write the PNG file to disk. | Provide a complete example that adds a pivot table, attaches a slicer, adjusts the print area to the slicer, renders the slicer image, and saves both the image and the workbook.
// Common Searches: how to set print area to slicer location using Aspose.Cells C# | export only slicer to PNG with Aspose.Cells .NET | retrieve slicer shape coordinates Aspose.Cells example | Aspose.Cells ImageOrPrintOptions OnlyArea slicer rendering | save workbook after adjusting print area for slicer Aspose.Cells
// Tags: set worksheet print area slicer bounds Aspose.Cells | render slicer to PNG image Aspose.Cells | create range from slicer shape Aspose.Cells | use ImageOrPrintOptions OnlyArea Aspose.Cells | pivot table slicer export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot; // Required for PivotTable related types

namespace SlicerPrintAreaRenderDemo
{
    // The sample creates a workbook, adds sample data, builds a pivot table, inserts a slicer for the pivot field, extracts the slicer's shape coordinates, sets the worksheet's print area to exactly those bounds, defines a range covering the slicer, renders that range to a PNG image using OnlyArea mode, saves the image, and finally saves the workbook with the updated print area.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a pivot table
                sheet.Cells["A1"].Value = "Fruit";
                sheet.Cells["B1"].Value = "Year";
                sheet.Cells["C1"].Value = "Amount";

                sheet.Cells["A2"].Value = "Apple";
                sheet.Cells["B2"].Value = 2020;
                sheet.Cells["C2"].Value = 50;

                sheet.Cells["A3"].Value = "Apple";
                sheet.Cells["B3"].Value = 2021;
                sheet.Cells["C3"].Value = 70;

                sheet.Cells["A4"].Value = "Banana";
                sheet.Cells["B4"].Value = 2020;
                sheet.Cells["C4"].Value = 30;

                sheet.Cells["A5"].Value = "Banana";
                sheet.Cells["B5"].Value = 2021;
                sheet.Cells["C5"].Value = 60;

                // Add a pivot table based on the data range
                int pivotIdx = sheet.PivotTables.Add("A1:C5", "E2", "FruitPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
                pivot.AddFieldToArea(PivotFieldType.Column, "Year");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer for the "Fruit" field of the pivot table
                int slicerIdx = sheet.Slicers.Add(pivot, "Fruit", "F2");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // Obtain the slicer's shape to determine its bounds
                var slicerShape = slicer.Shape;

                int ulRow = slicerShape.UpperLeftRow;      // zero‑based index
                int ulCol = slicerShape.UpperLeftColumn;
                int lrRow = slicerShape.LowerRightRow;
                int lrCol = slicerShape.LowerRightColumn;

                // Set the worksheet's print area to exactly the slicer's bounds
                string startCell = sheet.Cells[ulRow, ulCol].Name;
                string endCell   = sheet.Cells[lrRow, lrCol].Name;
                sheet.PageSetup.PrintArea = $"{startCell}:{endCell}";

                // Prepare image options – render only the defined area
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    OnlyArea = true
                };

                // Create a range that covers the slicer shape
                int rows = lrRow - ulRow + 1;
                int cols = lrCol - ulCol + 1;
                Aspose.Cells.Range slicerRange = sheet.Cells.CreateRange(ulRow, ulCol, rows, cols);

                // Convert the range (i.e., the slicer) to an image byte array
                byte[] imageData = slicerRange.ToImage(imgOptions);

                // Save the image to a file
                string outputPath = "SlicerImage.png";
                File.WriteAllBytes(outputPath, imageData);

                // Optionally, save the workbook to verify the print area
                string workbookPath = "WorkbookWithSlicerPrintArea.xlsx";
                workbook.Save(workbookPath);

                Console.WriteLine($"Slicer rendered to image: {outputPath}");
                Console.WriteLine($"Workbook saved with slicer print area: {workbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
