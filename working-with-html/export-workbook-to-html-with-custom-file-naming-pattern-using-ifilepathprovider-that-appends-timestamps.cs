// Title: Export Aspose.Cells Workbook to HTML with Timestamped Sheet Files via IFilePathProvider (C#)
// Description: Demonstrates how to save a multi‑sheet workbook as separate HTML files using Aspose.Cells. A custom IFilePathProvider appends a yyyyMMdd_HHmmss timestamp to each worksheet's file name, producing unique outputs such as FirstSheet_20230726_153045.html.
// Keywords: Aspose.Cells HTML export | IFilePathProvider custom naming | timestamped HTML files | C# Aspose.Cells save options | per‑sheet HTML export | .NET workbook to HTML | unique file names Aspose.Cells | HtmlSaveOptions timestamp
// Common Searches: Aspose.Cells IFilePathProvider example | C# export each worksheet to HTML with timestamp | custom file naming for Aspose.Cells HTML export | how to generate unique HTML file names per sheet Aspose.Cells | timestamped HTML output Aspose.Cells C#
// Developer Intent: Generate individual HTML files for each worksheet, embedding the sheet name and a timestamp in the file name.
// Use Cases: Automated reporting where each export must be versioned without overwriting previous files. | Web publishing of workbook data where distinct, time‑stamped HTML pages prevent cache conflicts. | Archiving daily snapshots of multi‑sheet workbooks for audit trails or compliance.
// AI Prompts: Create a C# IFilePathProvider that adds a GUID before the sheet name in each HTML file. | Modify the TimestampFilePathProvider to include the workbook's base name and a custom output folder. | Show how to combine HtmlSaveOptions with a custom IFilePathProvider to save HTML files to Azure Blob Storage.

using System;
using Aspose.Cells;

namespace AsposeCellsExportHtml
{
    // Demonstrates how to save a multi‑sheet workbook as separate HTML files using Aspose.Cells. A custom IFilePathProvider appends a yyyyMMdd_HHmmss timestamp to each worksheet's file name, producing unique outputs such as FirstSheet_20230726_153045.html.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();

            // First worksheet
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            sheet1.Cells["A1"].PutValue("Hello from Sheet1");

            // Second worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            sheet2.Cells["A1"].PutValue("Hello from Sheet2");

            // Configure HTML save options and assign the custom file path provider
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.FilePathProvider = new TimestampFilePathProvider();

            // Save the workbook as HTML; the provider will generate per‑sheet file names with timestamps
            workbook.Save("Workbook.html", saveOptions);
        }
    }

    // Custom IFilePathProvider that appends a timestamp to each worksheet's HTML file name
    public class TimestampFilePathProvider : IFilePathProvider
    {
        public string GetFullName(string sheetName)
        {
            // Timestamp format: yyyyMMdd_HHmmss
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            // Resulting file name, e.g., Sheet1_20230726_153045.html
            return $"{sheetName}_{timestamp}.html";
        }
    }
}
