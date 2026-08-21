// Title: Aspose.Cells for .NET – Embed Multiple Excel Worksheets as Separate PDF Attachments (C#)
// Description: Learn how to programmatically embed each worksheet of an Excel workbook as an individual OLE attachment inside a PDF using Aspose.Cells. The example creates a summary sheet, saves other sheets as temporary .xlsx files, adds them as icons, and generates a PDF with the EmbedAttachments option enabled. Includes cleanup of temporary files and cross‑platform guidance for developers in the US, UK, India and other regions.
// Keywords: Aspose.Cells embed worksheets PDF | C# PDFSaveOptions EmbedAttachments | Excel OLE objects in PDF | attach Excel sheets to PDF Aspose | .NET generate PDF with worksheet attachments | temporary file cleanup Aspose.Cells | summary sheet PDF icons
// Common Searches: How to add Excel worksheets as PDF attachments with Aspose.Cells C# | Aspose.Cells PDFSaveOptions EmbedAttachments example | Create PDF with OLE icons for each worksheet | Save multiple worksheets as separate files inside a PDF | C# embed Excel sheets in PDF using Aspose
// Developer Intent: Produce a single PDF where every non‑summary worksheet is stored as an individual attachment accessible via icons on the first sheet.
// Use Cases: Audit‑ready reports: a one‑page summary PDF with detailed data sheets attached for reviewers. | Client deliverables: distribute a compact PDF while still providing raw Excel data for analysis. | Regulatory archiving: store the full workbook inside a PDF, preserving original worksheets as attachments.
// AI Prompts: Generate C# code with Aspose.Cells that embeds all worksheets except the first one as OLE objects and saves the workbook as a PDF with embedded attachments. | Explain the role of PdfSaveOptions.EmbedAttachments and required OLE settings when converting Excel to PDF using Aspose.Cells. | Provide best‑practice recommendations for handling dozens of worksheets, temporary file management, and performance when embedding them as PDF attachments.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Learn how to programmatically embed each worksheet of an Excel workbook as an individual OLE attachment inside a PDF using Aspose.Cells. The example creates a summary sheet, saves other sheets as temporary .xlsx files, adds them as icons, and generates a PDF with the EmbedAttachments option enabled. Includes cleanup of temporary files and cross‑platform guidance for developers in the US, UK, India and other regions.
    public class EmbedMultipleWorksheetsAsAttachments
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and add three worksheets
            Workbook workbook = new Workbook();
            Worksheet mainSheet = workbook.Worksheets[0];
            mainSheet.Name = "Summary";

            // Worksheet 1
            Worksheet sheet1 = workbook.Worksheets.Add("SalesData");
            sheet1.Cells["A1"].PutValue("Product");
            sheet1.Cells["B1"].PutValue("Quantity");
            sheet1.Cells["A2"].PutValue("Apple");
            sheet1.Cells["B2"].PutValue(120);
            sheet1.Cells["A3"].PutValue("Banana");
            sheet1.Cells["B3"].PutValue(85);

            // Worksheet 2
            Worksheet sheet2 = workbook.Worksheets.Add("EmployeeInfo");
            sheet2.Cells["A1"].PutValue("Name");
            sheet2.Cells["B1"].PutValue("Department");
            sheet2.Cells["A2"].PutValue("John Doe");
            sheet2.Cells["B2"].PutValue("Finance");
            sheet2.Cells["A3"].PutValue("Jane Smith");
            sheet2.Cells["B3"].PutValue("HR");

            // Temporary folder for intermediate files
            string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeSheets");
            Directory.CreateDirectory(tempFolder);

            // Embed each worksheet (except the main summary sheet) as an OLE object into the main sheet
            int oleIndex = 0;
            for (int i = 1; i < workbook.Worksheets.Count; i++) // start from 1 to skip the summary sheet
            {
                Worksheet ws = workbook.Worksheets[i];
                string tempFile = Path.Combine(tempFolder, $"{ws.Name}.xlsx");

                // Save the single worksheet as a temporary Excel file
                ws.Workbook.Save(tempFile, SaveFormat.Xlsx);

                // Ensure the temporary file exists before embedding
                if (File.Exists(tempFile))
                {
                    // Position the icons vertically with some spacing
                    int row = 2 + oleIndex * 5;
                    // Add OLE object and get its index
                    oleIndex = mainSheet.OleObjects.Add(row, 1, 200, 50, File.ReadAllBytes(tempFile));
                    // Retrieve the OLE object to set its properties
                    OleObject ole = mainSheet.OleObjects[oleIndex];
                    ole.FileFormatType = FileFormatType.Xlsx;
                    ole.DisplayAsIcon = true;
                }

                // Clean up the temporary file after embedding
                try
                {
                    if (File.Exists(tempFile))
                        File.Delete(tempFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Unable to delete temporary file '{tempFile}'. {ex.Message}");
                }
            }

            // Prepare PDF save options to embed OLE attachments
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = true
            };

            // Save the workbook as PDF; the main sheet will contain icons for each embedded worksheet
            string outputPdf = Path.Combine(Environment.CurrentDirectory, "WorkbookWithAttachments.pdf");
            try
            {
                workbook.Save(outputPdf, pdfOptions);
                Console.WriteLine($"PDF created with embedded worksheet attachments: {outputPdf}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving PDF: {ex.Message}");
            }

            // Optional: remove the temporary folder
            try
            {
                if (Directory.Exists(tempFolder))
                    Directory.Delete(tempFolder, true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to delete temporary folder '{tempFolder}'. {ex.Message}");
            }
        }
    }
}
