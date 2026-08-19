// Title: Read the Language Built‑In Property of an Excel Workbook using Aspose.Cells for .NET
// Description: C# example that loads an Excel file with Aspose.Cells, accesses the BuiltInDocumentPropertyCollection, retrieves the Language property (the workbook's locale), and outputs it to the console.
// Keywords: Aspose.Cells read Language property | C# Excel built‑in document properties | retrieve workbook locale Aspose | load workbook with LoadOptions .NET | Excel language built‑in property
// Common Searches: How to get the Language built‑in property from an Excel file with Aspose.Cells | Aspose.Cells C# read document locale | Read Excel workbook language setting using Aspose | Get built‑in document properties Aspose.Cells .NET
// Developer Intent: Load an Excel workbook and obtain its Language built‑in property to determine the file’s locale.
// Use Cases: Detect the workbook’s locale to apply culture‑specific formatting or calculations. | Log or display the document language for auditing, reporting, or compliance. | Route processing logic based on the language setting of incoming Excel files.
// AI Prompts: Show how to read other built‑in properties such as Author, Title, and CreatedDate with Aspose.Cells. | Provide code to modify the Language property of a workbook and save the changes. | Explain how to handle cases where the Language property is missing or empty.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

// C# example that loads an Excel file with Aspose.Cells, accesses the BuiltInDocumentPropertyCollection, retrieves the Language property (the workbook's locale), and outputs it to the console.
class Program
{
    static void Main()
    {
        // Create LoadOptions using the default constructor (rule-provided)
        LoadOptions loadOptions = new LoadOptions();

        // Load the workbook with the specified LoadOptions
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Access the built‑in document properties collection
        BuiltInDocumentPropertyCollection builtInProps = workbook.BuiltInDocumentProperties;

        // Read the Language property which indicates the locale of the document
        string language = builtInProps.Language;

        Console.WriteLine($"Document language: {language}");
    }
}
