// Title: Batch convert Excel files to PDF with a PivotTable timeline using Aspose.Cells for .NET
// AI Prompts: Generate a C# console program that iterates over all .xlsx files in a directory, inserts a timeline based on the first PivotTable in each worksheet, and saves each workbook as a PDF with Aspose.Cells. | Write code that creates a temporary copy of each modified workbook, uses Aspose.Cells ConversionUtility to export it to PDF, and then removes the temporary file. | Add logging that prints the source Excel file path and the destination PDF path for every conversion performed.
// Common Searches: how to add a timeline to a pivot table with Aspose.Cells before PDF conversion | C# batch convert multiple Excel workbooks to PDF and insert timelines using Aspose.Cells | Aspose.Cells convert a folder of .xlsx files to PDF while automatically adding a timeline | example code for inserting a timeline into the first PivotTable and exporting to PDF in .NET
// Tags: batch Excel to PDF conversion with Aspose.Cells | insert pivot table timeline C# Aspose.Cells | ConversionUtility PDF export from workbook | temporary file cleanup Aspose.Cells | automated PDF generation with timelines

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Utility;

namespace BatchConvertWithTimeline
{
    // The C# console app scans a given input folder for .xlsx files, loads each workbook, adds a timeline to the first PivotTable on the first worksheet, saves a temporary copy, converts the temporary workbook to PDF using Aspose.Cells ConversionUtility, deletes the temporary file, and writes the resulting PDFs to a specified output directory while logging each operation.
    class Program
    {
        static void Main()
        {
            // Input folder containing Excel files
            string inputFolder = @"C:\InputExcel";
            // Output folder where PDFs will be saved
            string outputFolder = @"C:\OutputPdf";

            Directory.CreateDirectory(outputFolder);

            // Process each Excel file in the input folder
            foreach (string excelPath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                // Load the workbook
                Workbook workbook = new Workbook(excelPath);
                Worksheet sheet = workbook.Worksheets[0];

                // Add a Timeline if a PivotTable exists
                if (sheet.PivotTables.Count > 0)
                {
                    // Use the first PivotTable as the data source
                    PivotTable pivot = sheet.PivotTables[0];

                    // Use the first BaseField name for the Timeline
                    string baseFieldName = pivot.BaseFields[0].Name;

                    // Add the Timeline at cell E1 (you can change the address as needed)
                    sheet.Timelines.Add(pivot, "E1", baseFieldName);
                }

                // Save the modified workbook to a temporary file
                string tempPath = Path.Combine(Path.GetTempPath(),
                    Guid.NewGuid().ToString() + ".xlsx");
                workbook.Save(tempPath, SaveFormat.Xlsx);

                // Define the PDF output path (same file name, .pdf extension)
                string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Convert the temporary Excel file to PDF
                ConversionUtility.Convert(tempPath, pdfPath);

                // Clean up the temporary file
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }

                Console.WriteLine($"Converted '{excelPath}' to PDF with Timeline: '{pdfPath}'");
            }
        }
    }
}
