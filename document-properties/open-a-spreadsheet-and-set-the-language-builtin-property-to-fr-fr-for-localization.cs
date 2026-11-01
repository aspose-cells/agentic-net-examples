// Title: Set the Language built‑in property of an Excel workbook to French (fr-FR) with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set workbook built‑in language to fr-FR | Change Excel file locale to French using Aspose.Cells .NET | How to modify the Language document property in an .xlsx with Aspose.Cells | Programmatically set language property for Excel workbook in C# | Localize an Excel workbook by updating the Language built‑in property with Aspose.Cells
// Tags: Aspose.Cells workbook language property | C# Excel built‑in document properties | set language property fr-FR | Excel file localization Aspose.Cells | save workbook after property update

using Aspose.Cells;
using System;

// The example loads an existing Excel file (input.xlsx) into an Aspose.Cells Workbook, assigns the built‑in Language property the value "fr-FR" to apply French (France) localization, and saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Load the existing spreadsheet
        Workbook workbook = new Workbook("input.xlsx");

        // Set the Language built‑in property to French (France) for localization
        workbook.BuiltInDocumentProperties.Language = "fr-FR";

        // Save the modified spreadsheet
        workbook.Save("output.xlsx");
    }
}
