using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    public class EmbedWorksheetsAsAttachmentsInPdf
    {
        // Entry point for the console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and add sample worksheets
                Workbook workbook = new Workbook();
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";
                sheet1.Cells["A1"].PutValue("Data in Sheet 1");
                sheet1.Cells["A2"].PutValue(123);

                Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
                sheet2.Cells["A1"].PutValue("Data in Sheet 2");
                sheet2.Cells["A2"].PutValue(456);

                Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");
                sheet3.Cells["A1"].PutValue("Data in Sheet 3");
                sheet3.Cells["A2"].PutValue(789);

                // Directory for temporary worksheet files
                string tempDir = Path.Combine(Path.GetTempPath(), "AsposeSheetsTemp");
                Directory.CreateDirectory(tempDir);

                // Sheet that will hold the OLE attachments
                Worksheet attachmentSheet = workbook.Worksheets.Add("Attachments");
                int oleRow = 0; // starting row for OLE objects

                // Iterate through each worksheet (excluding the attachment sheet)
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet ws = workbook.Worksheets[i];
                    if (ws.Name == "Attachments")
                        continue; // skip the attachment sheet

                    // Save the current worksheet to a temporary Excel file
                    string tempFilePath = Path.Combine(tempDir, $"{ws.Name}.xlsx");
                    Workbook tempWb = new Workbook();
                    tempWb.Worksheets.AddCopy(ws.Name);
                    // Remove the default empty sheet that comes with a new workbook
                    if (tempWb.Worksheets.Count > 1)
                        tempWb.Worksheets.RemoveAt(0);
                    tempWb.Save(tempFilePath);

                    // Ensure the temporary file exists before reading
                    if (!File.Exists(tempFilePath))
                        continue;

                    // Add the temporary file as an OLE object (Excel attachment) to the attachment sheet
                    byte[] fileBytes = File.ReadAllBytes(tempFilePath);
                    int oleIndex = attachmentSheet.OleObjects.Add(oleRow, 0, 200, 200, fileBytes);
                    OleObject ole = attachmentSheet.OleObjects[oleIndex];
                    ole.FileFormatType = FileFormatType.Xlsx; // specify the format of the embedded file
                    ole.DisplayAsIcon = true;                 // show as an icon
                    // Note: Aspose.Cells OleObject does not expose IconCaption; the sheet name will appear as the default caption

                    // Move to next row for the next attachment
                    oleRow += 10;
                }

                // Configure PDF save options to embed OLE attachments
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    EmbedAttachments = true
                };

                // Save the workbook as PDF; the attachments will be embedded in the PDF
                string outputPdf = "WorkbookWithEmbeddedSheets.pdf";
                workbook.Save(outputPdf, pdfOptions);
                Console.WriteLine($"PDF saved with embedded worksheet attachments: {outputPdf}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during processing: {ex.Message}");
            }
            finally
            {
                // Clean up temporary files and directory
                try
                {
                    string tempDir = Path.Combine(Path.GetTempPath(), "AsposeSheetsTemp");
                    if (Directory.Exists(tempDir))
                        Directory.Delete(tempDir, true);
                }
                catch
                {
                    // If deletion fails, ignore – temporary files are not critical
                }
            }
        }
    }
}