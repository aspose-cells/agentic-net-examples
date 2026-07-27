using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Utility;

class RemoveSlicerAndConvertToPdf
{
    // Name of the slicer to be removed from each workbook
    const string TargetSlicerName = "MySlicer";

    // Folder containing the source XLSX files
    const string InputFolder = @"C:\InputWorkbooks";

    static void Main()
    {
        // Get all XLSX files in the input folder
        string[] workbookFiles = Directory.GetFiles(InputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string workbookPath in workbookFiles)
        {
            // Load the workbook
            Workbook wb = new Workbook(workbookPath);

            // Iterate through all worksheets
            foreach (Worksheet ws in wb.Worksheets)
            {
                SlicerCollection slicers = ws.Slicers;

                // Scan the slicer collection in reverse order and remove the matching slicer
                for (int i = slicers.Count - 1; i >= 0; i--)
                {
                    Slicer slicer = slicers[i];
                    // The Slicer class has a Name property that identifies the slicer
                    if (slicer.Name.Equals(TargetSlicerName, StringComparison.OrdinalIgnoreCase))
                    {
                        slicers.RemoveAt(i);
                    }
                }
            }

            // Save the modified workbook (overwrites the original file)
            wb.Save(workbookPath, SaveFormat.Xlsx);

            // Convert the updated workbook to PDF
            string pdfPath = Path.ChangeExtension(workbookPath, ".pdf");
            ConversionUtility.Convert(workbookPath, pdfPath);

            Console.WriteLine($"Processed '{Path.GetFileName(workbookPath)}' -> '{Path.GetFileName(pdfPath)}'");
        }

        Console.WriteLine("Batch processing completed.");
    }
}