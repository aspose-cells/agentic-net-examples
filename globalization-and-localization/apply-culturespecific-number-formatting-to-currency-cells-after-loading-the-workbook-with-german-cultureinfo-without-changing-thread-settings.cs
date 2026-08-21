// Title: Apply German‑Locale Currency Formatting to Excel Cells with Aspose.Cells for .NET (No Thread Culture Change)
// Description: Creates a sample workbook, loads it using LoadOptions.CultureInfo set to de‑DE, synchronizes wb.Settings.CultureInfo, defines a reusable currency style (Number = 5) that respects the workbook's locale, applies the style to every numeric cell, and saves the file so values display German currency symbols and comma decimal separators—all without altering the application thread culture.
// Keywords: Aspose.Cells | C# | .NET | German locale | de-DE | currency formatting | Excel localization | LoadOptions CultureInfo | Settings.CultureInfo | number format 5 | apply style to numeric cells | globalization | Excel workbook formatting
// Common Searches: Aspose.Cells German currency format example | Load Excel file with specific CultureInfo in .NET | Apply currency style to all numeric cells Aspose.Cells | Set workbook Settings.CultureInfo without changing thread culture | Number format 5 currency Aspose.Cells
// Developer Intent: Format all numeric cells as German currency after loading a workbook, using workbook‑level culture settings instead of modifying the thread culture.
// Use Cases: Generate a workbook, then reload it with LoadOptions.CultureInfo = new CultureInfo("de-DE") for German formatting. | Synchronize wb.Settings.CultureInfo with the load options to ensure style rendering follows the German locale. | Create a single currency style (Number = 5) and assign it to each cell where CellValueType.IsNumeric is true. | Save the workbook so that numbers appear with € symbol and comma as decimal separator in any viewer.
// AI Prompts: Write C# code that loads an Excel file using Aspose.Cells LoadOptions with de‑DE CultureInfo and applies the built‑in currency style (Number = 5) to all numeric cells. | Explain why Aspose.Cells uses the workbook's Settings.CultureInfo for number formatting and how this avoids the need to change Thread.CurrentThread.CurrentCulture.

using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsCultureSpecificCurrency
{
    // Creates a sample workbook, loads it using LoadOptions.CultureInfo set to de‑DE, synchronizes wb.Settings.CultureInfo, defines a reusable currency style (Number = 5) that respects the workbook's locale, applies the style to every numeric cell, and saves the file so values display German currency symbols and comma decimal separators—all without altering the application thread culture.
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a sample workbook with numeric values (for demo purposes)
            // -----------------------------------------------------------------
            string samplePath = "sample.xlsx";
            Workbook creator = new Workbook();
            Worksheet sheet = creator.Worksheets[0];

            // Populate some cells with numeric values
            sheet.Cells["A1"].PutValue(1234.56);   // will become currency
            sheet.Cells["A2"].PutValue(9876.54);   // will become currency
            sheet.Cells["B1"].PutValue(0.1234);    // non‑currency example

            // Save the sample workbook
            creator.Save(samplePath, SaveFormat.Xlsx);

            // ---------------------------------------------------------------
            // 2. Load the workbook using German culture (de-DE) via LoadOptions
            // ---------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.CultureInfo = new CultureInfo("de-DE"); // German uses comma as decimal separator

            Workbook wb = new Workbook(samplePath, loadOptions);

            // Ensure the workbook's own culture is also set (affects style formatting)
            wb.Settings.CultureInfo = new CultureInfo("de-DE");

            // ---------------------------------------------------------------
            // 3. Apply currency number format to cells that contain numeric values
            // ---------------------------------------------------------------
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Create a reusable currency style (built‑in number format 5 = Currency)
            Style currencyStyle = wb.CreateStyle();
            currencyStyle.Number = 5; // Currency format respects workbook's CultureInfo

            // Iterate through used cells and apply the currency style where appropriate
            foreach (Cell cell in cells)
            {
                // Check if the cell contains a numeric value (double, decimal, int, etc.)
                if (cell.Type == CellValueType.IsNumeric)
                {
                    // Apply the currency style
                    cell.SetStyle(currencyStyle);
                }
            }

            // ---------------------------------------------------------------
            // 4. Save the workbook with culture‑specific currency formatting
            // ---------------------------------------------------------------
            string outputPath = "output.xlsx";
            wb.Save(outputPath, SaveFormat.Xlsx);

            // Inform the user
            Console.WriteLine($"Workbook saved to '{outputPath}' with German currency formatting.");
        }
    }
}
