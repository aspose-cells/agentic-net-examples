// Title: Set Excel Workbook Language Property to French (fr-FR) with Aspose.Cells for .NET
// Description: Load an existing .xlsx file using Aspose.Cells, assign the built‑in Language property the culture code "fr-FR" to localize the workbook for French (France), and save the updated file.
// Keywords: Aspose.Cells set language | Excel built‑in Language property .NET | fr-FR workbook locale | localize Excel file Aspose | document properties Aspose.Cells
// Common Searches: Aspose.Cells change workbook language to French | set built‑in Language property C# Aspose.Cells | how to localize Excel file with Aspose.Cells | update Excel document locale programmatically
// Developer Intent: Programmatically set the built‑in Language property of an Excel workbook to "fr-FR" and persist the change.
// Use Cases: Prepare a French‑language template for distribution to French‑speaking users. | Ensure generated reports comply with regional locale settings before archiving. | Automate localization of existing spreadsheets in a batch processing pipeline.
// AI Prompts: Generate C# code that uses Aspose.Cells to set the Language built‑in property of a workbook to any given culture code and save the file. | Explain the impact of the Language document property on Excel UI language and how to verify it after modification with Aspose.Cells. | Show a single Aspose.Cells workflow that updates multiple built‑in properties such as Language, Author, and Title in one pass.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// Load an existing .xlsx file using Aspose.Cells, assign the built‑in Language property the culture code "fr-FR" to localize the workbook for French (France), and save the updated file.
class SetLanguageDemo
{
    static void Main()
    {
        // Load the existing spreadsheet
        Workbook workbook = new Workbook("input.xlsx"); // replace with your file path

        // Set the built‑in Language property to French (France)
        workbook.BuiltInDocumentProperties.Language = "fr-FR";

        // Save the workbook with the updated property
        workbook.Save("output.xlsx"); // replace with desired output path
    }
}
