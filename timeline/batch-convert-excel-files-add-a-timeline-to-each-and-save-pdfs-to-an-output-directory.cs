using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;
using Aspose.Cells.Utility;

class BatchConvertWithTimeline
{
    static void Main()
    {
        // Input folder containing Excel files
        string inputFolder = @"C:\InputExcel";
        // Output folder where PDFs will be saved
        string outputFolder = @"C:\OutputPdf";

        try
        {
            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each .xlsx file in the input folder
            foreach (string excelPath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                // Verify the file still exists before loading
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"File not found (skipped): {excelPath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(excelPath);
                    Worksheet sheet = workbook.Worksheets[0];

                    // ------------------------------------------------------------
                    // Create a simple PivotTable as a data source for the Timeline.
                    // This example assumes that the source data occupies A1:B10.
                    // Adjust the range as needed for real data.
                    // ------------------------------------------------------------
                    int pivotIndex = sheet.PivotTables.Add("A1:B10", "D1", "PivotTable1");
                    PivotTable pivot = sheet.PivotTables[pivotIndex];

                    // Add the first column (index 0) as a Row field – typically a date field.
                    pivot.AddFieldToArea(PivotFieldType.Row, 0);

                    // ------------------------------------------------------------
                    // Add a Timeline control linked to the PivotTable.
                    // The Timeline will be placed with its upper‑left corner at cell E1
                    // and will use the field named "Date" from the PivotTable.
                    // ------------------------------------------------------------
                    sheet.Timelines.Add(pivot, "E1", "Date");

                    // Save the modified workbook to a temporary file (required before conversion).
                    string tempExcelPath = Path.Combine(outputFolder,
                        Path.GetFileNameWithoutExtension(excelPath) + "_temp.xlsx");
                    workbook.Save(tempExcelPath);

                    // Convert the temporary Excel file to PDF using the provided utility.
                    string pdfPath = Path.Combine(outputFolder,
                        Path.GetFileNameWithoutExtension(excelPath) + ".pdf");
                    ConversionUtility.Convert(tempExcelPath, pdfPath);

                    // Optional: delete the temporary Excel file.
                    File.Delete(tempExcelPath);

                    Console.WriteLine($"Converted '{excelPath}' to PDF with Timeline: '{pdfPath}'");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}