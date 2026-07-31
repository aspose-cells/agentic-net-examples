// Title: Batch Hide Empty Worksheets in Excel Workbooks with Aspose.Cells for .NET (C#)
// Description: A C# console utility that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, evaluates every worksheet using MaxDataRow/MaxDataColumn, hides sheets that contain no data (IsVisible = false), and saves the cleaned workbooks to a target directory. Perfect for bulk Excel cleanup.
// Keywords: Aspose.Cells | C# | hide empty worksheets | batch Excel processing | MaxDataRow | MaxDataColumn | programmatic sheet visibility | .NET Excel automation | Excel workbook cleanup | folder based workbook iteration
// Common Searches: C# hide empty Excel sheets Aspose.Cells | batch hide blank worksheets .NET | detect empty worksheet Aspose.Cells | process multiple workbooks Aspose.Cells C# | automate Excel sheet visibility Aspose.Cells
// Developer Intent: Automatically hide any worksheet that lacks data in each workbook and write the updated files back to disk.
// Use Cases: Remove blank tabs from generated reports before distribution to keep the workbook tidy. | Pre‑process a batch of user‑uploaded Excel files, hiding empty worksheets to reduce visual clutter in a document management system. | Automate cleanup of archived workbooks so that only populated worksheets remain visible.
// AI Prompts: Generate C# code using Aspose.Cells that hides empty worksheets in a single workbook and saves the result. | Refactor the program to log the names of hidden sheets and add support for both .xlsx and .xls formats. | Explain how MaxDataRow and MaxDataColumn can be used to determine whether a worksheet is empty in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace HideEmptyWorksheetsDemo
{
    // A C# console utility that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, evaluates every worksheet using MaxDataRow/MaxDataColumn, hides sheets that contain no data (IsVisible = false), and saves the cleaned workbooks to a target directory. Perfect for bulk Excel cleanup.
    class Program
    {
        static void Main()
        {
            // Folder containing the workbooks to process
            string inputFolder = @"C:\InputWorkbooks";
            // Folder where the processed workbooks will be saved (can be the same as input)
            string outputFolder = @"C:\OutputWorkbooks";

            try
            {
                // Ensure input folder exists; if not, create it to avoid DirectoryNotFoundException
                if (!Directory.Exists(inputFolder))
                {
                    Directory.CreateDirectory(inputFolder);
                    Console.WriteLine($"Input folder created at '{inputFolder}'. Place Excel files there and rerun the program.");
                    return;
                }

                // Ensure output folder exists
                Directory.CreateDirectory(outputFolder);

                // Get all Excel files in the input folder
                string[] workbookFiles = Directory.GetFiles(inputFolder, "*.xlsx");

                foreach (string filePath in workbookFiles)
                {
                    try
                    {
                        // Verify the file still exists before loading
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"File not found: {filePath}");
                            continue;
                        }

                        // Load the workbook
                        Workbook workbook = new Workbook(filePath);

                        // Iterate over all worksheets
                        foreach (Worksheet sheet in workbook.Worksheets)
                        {
                            // Determine if the worksheet contains any data
                            bool hasData = sheet.Cells.MaxDataRow >= 0 && sheet.Cells.MaxDataColumn >= 0;

                            // Hide the worksheet if it has no data
                            if (!hasData)
                            {
                                sheet.IsVisible = false;
                            }
                        }

                        // Build the output file path (overwrite the original file if same folder)
                        string fileName = Path.GetFileName(filePath);
                        string outputPath = Path.Combine(outputFolder, fileName);

                        // Save the modified workbook
                        workbook.Save(outputPath, SaveFormat.Xlsx);
                        Console.WriteLine($"Processed and saved: {outputPath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Processing completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
