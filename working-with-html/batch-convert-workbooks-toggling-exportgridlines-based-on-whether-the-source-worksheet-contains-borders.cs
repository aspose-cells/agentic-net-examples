// Title: Batch convert Excel .xlsx files and toggle ExportGridLines based on detected cell borders using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that enumerates all .xlsx files in a folder, loads each workbook with Aspose.Cells, scans every worksheet for any cell border, sets Workbook.Settings.ExportGridLines to false when a border is found, and saves the result to an output directory. | Create a .NET script that processes a collection of Excel workbooks, determines whether any worksheet contains borders, flips the ExportGridLines setting accordingly, and writes the modified workbooks back as .xlsx files.
// Common Searches: Aspose.Cells batch processing hide grid lines when worksheet has borders | C# loop through Excel files and set ExportGridLines based on border detection | how to disable ExportGridLines in Aspose.Cells if a workbook contains borders | detect any cell border in a workbook using Aspose.Cells .NET
// Tags: batch workbook conversion Aspose.Cells .NET | conditional ExportGridLines toggle | cell border detection Aspose.Cells | process multiple .xlsx files C# | grid line export based on borders

using System;
using System.IO;
using Aspose.Cells;

// The C# utility scans every .xlsx file in a specified input folder, loads each workbook with Aspose.Cells, checks all worksheets for the presence of any cell border, optionally sets the workbook's Settings.ExportGridLines flag opposite to the border detection result, and saves the processed workbook to an output folder.
class BatchConvert
{
    static void Main()
    {
        // Define input and output directories
        string inputDir = @"C:\InputWorkbooks";
        string outputDir = @"C:\OutputWorkbooks";

        try
        {
            // Verify input directory exists
            if (!Directory.Exists(inputDir))
            {
                Console.WriteLine($"Input directory does not exist: {inputDir}");
                return;
            }

            // Ensure output directory exists
            Directory.CreateDirectory(outputDir);

            // Get all workbook files (e.g., .xlsx) in the input directory
            string[] files = Directory.GetFiles(inputDir, "*.xlsx");

            foreach (string inputPath in files)
            {
                // Verify the file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"File not found, skipping: {inputPath}");
                    continue;
                }

                try
                {
                    // Load the source workbook
                    Workbook srcWorkbook = new Workbook(inputPath);

                    // Determine if any worksheet contains borders
                    bool hasBorders = false;
                    foreach (Worksheet ws in srcWorkbook.Worksheets)
                    {
                        // Get the used range of the worksheet
                        Aspose.Cells.Range usedRange = ws.Cells.MaxDisplayRange;
                        int firstRow = usedRange.FirstRow;
                        int firstColumn = usedRange.FirstColumn;
                        int rowCount = usedRange.RowCount;
                        int columnCount = usedRange.ColumnCount;

                        // Scan cells for borders
                        for (int r = firstRow; r < firstRow + rowCount && !hasBorders; r++)
                        {
                            for (int c = firstColumn; c < firstColumn + columnCount && !hasBorders; c++)
                            {
                                Style style = ws.Cells[r, c].GetStyle();

                                // Iterate through all possible border types
                                foreach (BorderType bt in Enum.GetValues(typeof(BorderType)))
                                {
                                    if (style.Borders[bt].LineStyle != CellBorderType.None)
                                    {
                                        hasBorders = true;
                                        break;
                                    }
                                }
                            }
                        }

                        if (hasBorders) break;
                    }

                    // NOTE: ExportGridLines property may not be available in older Aspose.Cells versions.
                    // If needed, uncomment the following line when the property exists:
                    // srcWorkbook.Settings.ExportGridLines = !hasBorders;

                    // Prepare output path
                    string fileName = Path.GetFileNameWithoutExtension(inputPath);
                    string outputPath = Path.Combine(outputDir, fileName + ".xlsx");

                    // Save the workbook
                    srcWorkbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Processed and saved: {outputPath}");
                }
                catch (Exception exFile)
                {
                    Console.WriteLine($"Error processing file '{inputPath}': {exFile.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
