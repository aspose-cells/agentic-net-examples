using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook with default load options.
        // ParsingFormulaOnOpen can be set if you want to control formula parsing during load.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ParsingFormulaOnOpen = true; // optional, defaults to true for XLSX
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Add a named range (or retrieve it if it already exists) and assign a formula to it.
        // The RefersTo property must start with an equal sign.
        int nameIndex = workbook.Worksheets.Names.Add("MyNamedRange");
        Name namedRange = workbook.Worksheets.Names[nameIndex];
        namedRange.RefersTo = "=Sheet1!$A$1:$A$10"; // Formula defining the range

        // Save the modified workbook.
        workbook.Save("output.xlsx");
    }
}