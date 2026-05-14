using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class SetLanguagePropertyExample
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the built‑in document properties collection
        BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

        // Set the Language property to French (France)
        builtInProps.Language = "fr-FR";

        // Save the workbook with the updated language property
        workbook.Save("output.xlsx");
    }
}