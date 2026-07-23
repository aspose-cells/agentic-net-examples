// Title: Set Excel Language Built‑in Property to French (fr-FR) using Aspose.Cells for .NET (C#)
// Description: This C# example shows how to open an existing .xlsx file with Aspose.Cells, modify the BuiltInDocumentPropertyCollection to set the Language property to the locale code “fr‑FR”, and save the workbook. Updating the Language property ensures the file is recognized as French (France) in Excel and improves regional compatibility.
// Keywords: Aspose.Cells C# set language property | Excel built‑in document properties | fr-FR locale Excel | update workbook language .NET | localize Excel file metadata | Aspose.Cells document properties example | change Excel language programmatically
// Common Searches: Aspose.Cells set language property C# | How to change Excel file language to French programmatically | Update built‑in document properties in .NET | Set Language built‑in property Aspose.Cells | C# code to set Excel locale fr-FR
// Developer Intent: Programmatically set the workbook’s Language built‑in property to French (fr-FR) and save the file.
// Use Cases: Generate French‑language reports automatically by embedding the correct locale before distribution. | Batch‑process workbooks to comply with regional standards requiring the Language metadata to match the target market. | Ensure Excel’s language‑specific features (e.g., date formats, spell‑check) activate correctly for French users.
// AI Prompts: Provide C# code using Aspose.Cells that opens an .xlsx file, sets the Language built‑in property to a specified locale (e.g., fr-FR), and saves the updated workbook. | Create a reusable function `SetWorkbookLanguage(string inputPath, string locale, string outputPath)` that updates the Language property with Aspose.Cells. | Explain the impact of the Language built‑in property on Excel localization and how to verify the change with Aspose.Cells APIs.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// This C# example shows how to open an existing .xlsx file with Aspose.Cells, modify the BuiltInDocumentPropertyCollection to set the Language property to the locale code “fr‑FR”, and save the workbook. Updating the Language property ensures the file is recognized as French (France) in Excel and improves regional compatibility.
class SetLanguageProperty
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the built‑in document properties collection
        BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

        // Set the Language property to French (France)
        builtInProps.Language = "fr-FR";

        // Save the workbook with the updated property
        workbook.Save("output.xlsx");
    }
}
