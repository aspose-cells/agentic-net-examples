using System;
using Aspose.Cells;
using Aspose.Cells.Properties;
using Aspose.Cells.Rendering;

namespace AsposeCellsLanguageDemo
{
    class Program
    {
        static void Main()
        {
            // Scenario 1: Set UI language of the workbook via WorkbookSettings.LanguageCode
            SetWorkbookSettingsLanguage();

            // Scenario 2: Set document language via BuiltInDocumentProperties.Language
            SetBuiltInDocumentPropertyLanguage();

            // Scenario 3: Load a workbook with a specific language code using LoadOptions
            LoadWorkbookWithLanguageOption();

            // Scenario 4: Customize globalization settings (e.g., list separator) which affect language‑dependent formatting
            SetGlobalizationSettingsExample();

            // Scenario 5: Use DefaultEditLanguage in AutoFitterOptions for CJK text layout
            AutoFitWithDefaultEditLanguage();

            // Scenario 6: Set DefaultEditLanguage when saving to PDF to affect rendering
            SavePdfWithDefaultEditLanguage();
        }

        // Scenario 1: Set UI language of the workbook via WorkbookSettings.LanguageCode
        static void SetWorkbookSettingsLanguage()
        {
            Workbook wb = new Workbook();
            // Set the UI language to German (Germany)
            wb.Settings.LanguageCode = CountryCode.Germany;
            wb.Worksheets[0].Cells["A1"].PutValue("German UI");
            wb.Save("Workbook_LanguageCode_Germany.xlsx");
        }

        // Scenario 2: Set document language via BuiltInDocumentProperties.Language
        static void SetBuiltInDocumentPropertyLanguage()
        {
            Workbook wb = new Workbook();
            // Set the document language to French (France) using BCP‑47 tag
            wb.BuiltInDocumentProperties.Language = "fr-FR";
            wb.Worksheets[0].Cells["A1"].PutValue("Document language: French");
            wb.Save("Workbook_BuiltInLanguage_fr-FR.xlsx");
        }

        // Scenario 3: Load a workbook with a specific language code using LoadOptions
        static void LoadWorkbookWithLanguageOption()
        {
            // Load the workbook created in Scenario 1 with a different UI language for loading
            LoadOptions lo = new LoadOptions();
            lo.LanguageCode = CountryCode.USA; // Set UI language to English (USA) for loading
            Workbook wb = new Workbook("Workbook_LanguageCode_Germany.xlsx", lo);
            wb.Worksheets[0].Cells["A2"].PutValue("Loaded with USA language code");
            wb.Save("Workbook_LoadedWithUSALanguage.xlsx");
        }

        // Scenario 4: Customize globalization settings (e.g., list separator) which affect language‑dependent formatting
        static void SetGlobalizationSettingsExample()
        {
            Workbook wb = new Workbook();
            SettableGlobalizationSettings gs = new SettableGlobalizationSettings();

            // Use semicolon as list separator (common in many European locales)
            gs.SetListSeparator(';');

            // Override boolean display strings to German equivalents
            gs.SetBooleanValueString(true, "WAHR");
            gs.SetBooleanValueString(false, "FALSCH");

            // Apply the custom globalization settings to the workbook
            wb.Settings.GlobalizationSettings = gs;

            wb.Worksheets[0].Cells["A1"].PutValue(true);
            wb.Worksheets[0].Cells["A2"].PutValue(false);
            wb.Save("Workbook_CustomGlobalization.xlsx");
        }

        // Scenario 5: Use DefaultEditLanguage in AutoFitterOptions for CJK text layout
        static void AutoFitWithDefaultEditLanguage()
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Insert Japanese text to demonstrate CJK layout handling
            ws.Cells["A1"].PutValue("日本語のテキスト");

            // Configure AutoFitterOptions to treat the text as CJK
            AutoFitterOptions opts = new AutoFitterOptions();
            opts.DefaultEditLanguage = DefaultEditLanguage.CJK;

            // Auto‑fit rows using the specified language settings
            ws.AutoFitRows(opts);
            wb.Save("Workbook_AutoFit_CJK.xlsx");
        }

        // Scenario 6: Set DefaultEditLanguage when saving to PDF to affect rendering
        static void SavePdfWithDefaultEditLanguage()
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Insert Chinese text to demonstrate PDF rendering with CJK language
            ws.Cells["A1"].PutValue("中文文本");

            // Configure PDF save options to use CJK as the default edit language
            PdfSaveOptions pdfOpts = new PdfSaveOptions();
            pdfOpts.DefaultEditLanguage = DefaultEditLanguage.CJK;

            // Save the workbook as PDF with the language‑aware options
            wb.Save("Workbook_Pdf_CJK.pdf", pdfOpts);
        }
    }
}