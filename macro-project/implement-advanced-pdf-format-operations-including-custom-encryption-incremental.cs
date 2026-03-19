using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering.PdfSecurity;
using Aspose.Cells.Metadata;
using Aspose.Cells.Properties;

class AdvancedPdfOperations
{
    static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a workbook with sample data and save it as Excel.
        // ------------------------------------------------------------
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue("Advanced PDF Operations");
        ws.Cells["A2"].PutValue(DateTime.Now);
        ws.Cells["B1"].PutValue("Encryption Demo");
        ws.Cells["B2"].PutValue("Metadata Demo");

        string excelPath = "Sample.xlsx";
        wb.Save(excelPath); // create the base Excel file

        // ------------------------------------------------------------
        // 2. Load the workbook metadata, add/modify properties and encrypt it.
        // ------------------------------------------------------------
        // Combine Encryption and DocumentProperties flags.
        MetadataOptions encryptOptions = new MetadataOptions(MetadataType.Encryption | MetadataType.DocumentProperties);
        WorkbookMetadata meta = new WorkbookMetadata(excelPath, encryptOptions);

        // Built‑in properties
        BuiltInDocumentPropertyCollection builtIn = meta.BuiltInDocumentProperties;
        builtIn.Title = "Advanced PDF Demo";
        builtIn.Author = "John Doe";
        builtIn.Company = "Acme Corp";

        // Custom properties
        CustomDocumentPropertyCollection custom = meta.CustomDocumentProperties;
        custom.Add("Project", "PDF Security");
        custom.Add("Version", "1.0");

        // Save the encrypted workbook.
        string encryptedExcelPath = "SampleEncrypted.xlsx";
        meta.Save(encryptedExcelPath);

        // ------------------------------------------------------------
        // 3. Load the encrypted workbook (decryption is handled automatically).
        // ------------------------------------------------------------
        Workbook decryptedWb = new Workbook(encryptedExcelPath);

        // ------------------------------------------------------------
        // 4. Configure PDF security (custom encryption) and other options.
        // ------------------------------------------------------------
        PdfSecurityOptions pdfSec = new PdfSecurityOptions
        {
            OwnerPassword = "Owner@123",
            UserPassword = "User@123",
            PrintPermission = true,
            ModifyDocumentPermission = false,
            ExtractContentPermission = false,
            FillFormsPermission = true,
            AnnotationsPermission = true,
            AccessibilityExtractContent = true,
            AssembleDocumentPermission = false,
            FullQualityPrintPermission = true
        };

        PdfSaveOptions pdfOpts = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b,               // PDF/A‑1b compliance
            SecurityOptions = pdfSec,                       // Apply the security settings
            OptimizationType = PdfOptimizationType.MinimumSize,
            CreatedTime = DateTime.Now,
            Producer = "Aspose.Cells Advanced Demo"
        };

        // Add a watermark to illustrate detailed PDF customization.
        RenderingFont watermarkFont = new RenderingFont("Calibri", 68)
        {
            Italic = true,
            Bold = true,
            Color = Color.Blue
        };
        pdfOpts.Watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
        {
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            Rotation = 45,
            Opacity = 0.5f,
            ScaleToPagePercent = 50
        };

        // ------------------------------------------------------------
        // 5. Save the decrypted workbook as a secured PDF.
        // ------------------------------------------------------------
        string pdfPath = "AdvancedOutput.pdf";
        decryptedWb.Save(pdfPath, pdfOpts);

        // ------------------------------------------------------------
        // 6. Demonstrate an incremental update: reopen the PDF, add a page, and save again.
        //    (Aspose.Cells does not support direct PDF incremental updates, so we simulate
        //     it by adding a new worksheet and re‑exporting.)
        // ------------------------------------------------------------
        Workbook wbForUpdate = new Workbook(encryptedExcelPath); // Load again for update.
        Worksheet newWs = wbForUpdate.Worksheets.Add("Incremental");
        newWs.Cells["A1"].PutValue("Incremental Update");
        newWs.Cells["B1"].PutValue(DateTime.Now);

        // Re‑export to PDF using the same security options (simulating incremental save).
        string pdfIncrementalPath = "AdvancedOutput_Incremental.pdf";
        wbForUpdate.Save(pdfIncrementalPath, pdfOpts);
    }
}