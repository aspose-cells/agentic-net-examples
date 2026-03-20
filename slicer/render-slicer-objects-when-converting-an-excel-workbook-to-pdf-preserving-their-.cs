using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerPdfDemo
{
    class Program
    {
        static void Main()
        {
            // Determine the full path of the input workbook (assumed to be in the same folder as the executable)
            string inputFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "InputWithSlicers.xlsx");
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputFile);

            // Ensure that all drawing objects (including slicers) are displayed
            workbook.Settings.DisplayDrawingObjects = DisplayDrawingObjects.DisplayShapes;

            // Iterate through each worksheet and make slicers printable
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Slicer slicer in sheet.Slicers)
                {
                    slicer.IsPrintable = true;
                }
            }

            // Configure PDF save options (default options are sufficient for slicer rendering)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Determine the output PDF path
            string outputFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OutputWithSlicers.pdf");

            // Save the workbook as PDF, preserving slicer appearance
            workbook.Save(outputFile, pdfOptions);

            Console.WriteLine($"PDF saved successfully to: {outputFile}");
        }
    }
}