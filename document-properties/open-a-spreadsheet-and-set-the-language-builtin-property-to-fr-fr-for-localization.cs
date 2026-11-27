// Title: How to set the Built‑In Language property to French (fr‑FR) in an Excel workbook with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set BuiltInDocumentProperties.Language to fr-FR | change Excel workbook language metadata using Aspose.Cells .NET | localize an .xlsx file by updating the Language property in C# | Aspose.Cells example for setting workbook locale to French | programmatically modify Excel file language built‑in property with Aspose.Cells
// Tags: set workbook language property Aspose.Cells | built‑in document properties language C# | Excel file localization Aspose.Cells .NET | modify language property in .xlsx with Aspose.Cells | Excel file locale configuration Aspose.Cells

using System;
using Aspose.Cells;

// The code loads an existing workbook, assigns the French locale identifier "fr-FR" to the BuiltInDocumentProperties.Language property for localization, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Set the Language built‑in property for localization (e.g., French - France)
        workbook.BuiltInDocumentProperties.Language = "fr-FR";

        // Save the workbook with the updated property
        workbook.Save("output.xlsx");
    }
}
