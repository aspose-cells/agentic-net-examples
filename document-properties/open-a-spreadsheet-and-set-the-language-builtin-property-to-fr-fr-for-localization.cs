using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class SetLanguageDemo
{
    static void Main()
    {
        // Load an existing workbook (provide the correct path to your file)
        Workbook workbook = new Workbook("input.xlsx");

        // Set the built‑in Language property to French (France)
        workbook.BuiltInDocumentProperties.Language = "fr-FR";

        // Save the workbook with the updated language setting
        workbook.Save("output.xlsx");
    }
}