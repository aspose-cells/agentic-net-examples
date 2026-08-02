// Title: Save an Excel file with locale‑specific number formats using LightCellsDataProvider in Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, insert numeric values, set the workbook's CultureInfo to French (fr-FR) and Region to France, configure OoxmlSaveOptions with a custom LightCellsDataProvider that supplies no data, and save the file as CultureSpecificWorkbook.xlsx. The example shows that locale settings control decimal separators and other number formatting during a LightCells save.
// Keywords: Aspose.Cells | C# | .NET | LightCellsDataProvider | OoxmlSaveOptions | CultureInfo | locale | French culture | fr-FR | number format | region settings | high‑performance Excel export | Excel workbook culture | save workbook with locale
// Common Searches: Aspose.Cells set workbook culture before saving | How to apply French number format when exporting Excel with Aspose.Cells | LightCellsDataProvider custom save options example | Save Excel with locale‑specific formatting in .NET | Configure CultureInfo and Region for Aspose.Cells workbook
// Developer Intent: Apply specific CultureInfo and Region to a workbook and save it using a custom LightCellsDataProvider for high‑performance, locale‑aware Excel export.
// Use Cases: Generate a French‑localized financial report where decimals appear as commas. | Export large datasets quickly while preserving locale‑specific number formats. | Create multiple regional versions of the same workbook by switching CultureInfo (e.g., fr-FR, de-DE) before each save.
// AI Prompts: Show code to save an Excel workbook with German (de-DE) culture settings using LightCellsDataProvider in Aspose.Cells for .NET. | Provide a C# example that changes CultureInfo and Region for several locales before exporting to XLSX with Aspose.Cells. | Explain the interaction between LightCellsDataProvider and workbook culture settings during the save process.

using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsCultureSpecificSave
{
    // Custom LightCellsDataProvider that does not supply any data.
    // It is used here only to demonstrate setting the provider while
    // applying culture-specific settings to the workbook.
    // Demonstrates how to create a workbook, insert numeric values, set the workbook's CultureInfo to French (fr-FR) and Region to France, configure OoxmlSaveOptions with a custom LightCellsDataProvider that supplies no data, and save the file as CultureSpecificWorkbook.xlsx. The example shows that locale settings control decimal separators and other number formatting during a LightCells save.
    public class CustomLightCellsDataProvider : LightCellsDataProvider
    {
        // No string gathering is required.
        public bool IsGatherString() => false;

        // No sheets are processed by this provider.
        public int SheetCount => 0;

        // Return false to indicate that the sheet should be processed
        // by the default saving mechanism.
        public bool StartSheet(int sheetIndex) => false;

        // No rows to iterate.
        public int NextRow() => -1;

        // No special row handling.
        public void StartRow(Row row) { }

        // No cells to iterate.
        public int NextCell() => -1;

        // No special cell handling.
        public void StartCell(Cell cell) { }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and add sample numeric data.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(12345.67);
            sheet.Cells["A2"].PutValue(98765.43);

            // Apply culture-specific settings.
            // For example, French culture uses comma as decimal separator.
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");
            // Optionally set the region as well.
            workbook.Settings.Region = CountryCode.France;

            // Create OoxmlSaveOptions and assign the custom LightCellsDataProvider.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
            {
                LightCellsDataProvider = new CustomLightCellsDataProvider()
            };

            // Save the workbook using the save options.
            workbook.Save("CultureSpecificWorkbook.xlsx", saveOptions);
        }
    }
}
