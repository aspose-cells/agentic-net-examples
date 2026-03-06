using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class SetWorkbookLanguage
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Access the built‑in document properties collection
        BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

        // Set the language of the workbook (e.g., English - United States)
        builtInProps.Language = "en-US";

        // Additional properties for demonstration (optional)
        builtInProps.Author = "John Doe";
        builtInProps.Title = "Language Property Demo";

        // Save the workbook to disk
        workbook.Save("LanguageDemo.xlsx");
    }
}