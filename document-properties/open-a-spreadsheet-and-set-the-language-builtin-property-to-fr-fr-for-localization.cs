using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing spreadsheet
        Workbook workbook = new Workbook("input.xlsx");

        // Set the built‑in Language property to French (France)
        workbook.BuiltInDocumentProperties.Language = "fr-FR";

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}