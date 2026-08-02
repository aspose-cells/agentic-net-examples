// Title: Pass a Locale Identifier with SetVariable to Localize Smart Marker Formatting in Aspose.Cells for .NET
// Description: This example shows how to use WorkbookDesigner.SetVariable to supply a locale identifier (e.g., "fr-FR") that drives date and number formatting for smart markers. The workbook’s culture is set, smart markers are processed, and the result is saved as an Excel file with locale‑specific presentation.
// Keywords: Aspose.Cells SetVariable locale | smart markers localization .NET | Excel culture formatting Aspose | date number locale smart markers | WorkbookDesigner SetVariable example | multi‑language Excel reports Aspose | culture‑aware smart markers | Aspose.Cells locale identifier
// Common Searches: Aspose.Cells SetVariable locale identifier | how to localize smart marker dates in .NET | change number format for smart markers using culture | set French culture for Aspose.Cells smart markers | locale‑aware Excel export Aspose.Cells
// Developer Intent: Provide a locale identifier through SetVariable so that smart markers automatically format dates and numbers according to the specified culture.
// Use Cases: Generate a French‑formatted financial statement where dates appear as DD/MM/YYYY and decimals use commas. | Create a single invoice template that switches between German and US number formats by changing the locale variable before processing. | Build a dashboard that adapts its date and currency display to the user’s regional settings without manual formatting code.
// AI Prompts: Show how to call WorkbookDesigner.SetVariable with a culture code to affect smart marker formatting. | Explain the effect of Workbook.Settings.CultureInfo versus SetVariable on locale‑specific smart markers. | Provide a step‑by‑step guide to produce locale‑aware Excel files using Aspose.Cells smart markers and SetVariable.

using System;
using System.Data;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // This example shows how to use WorkbookDesigner.SetVariable to supply a locale identifier (e.g., "fr-FR") that drives date and number formatting for smart markers. The workbook’s culture is set, smart markers are processed, and the result is saved as an Excel file with locale‑specific presentation.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Build a simple template with smart markers
                // Header row
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Amount");

                // Smart marker row – values will be filled from the data source
                sheet.Cells["A2"].PutValue("&=Date");
                sheet.Cells["B2"].PutValue("&=Amount");

                // Prepare a DataTable that will serve as the data source for smart markers
                DataTable dt = new DataTable("Data");
                dt.Columns.Add("Date", typeof(DateTime));
                dt.Columns.Add("Amount", typeof(double));

                // Add sample rows
                dt.Rows.Add(new DateTime(2023, 1, 15), 1234.56);
                dt.Rows.Add(new DateTime(2023, 2, 20), 7890.12);
                dt.Rows.Add(new DateTime(2023, 3, 5), 345.67);

                // Create a WorkbookDesigner to process smart markers
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Set the data source (the table name matches the smart marker prefix, if any)
                designer.SetDataSource(dt);

                // Set the workbook's culture info to match the desired locale (e.g., French)
                workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

                // Process the smart markers – they will be replaced with formatted values
                designer.Process();

                // Save the workbook
                string outputPath = "SmartMarkerLocaleDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
