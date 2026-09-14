// Title: Remove named ranges that reference missing worksheets in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates the workbook's Names collection, identifies entries whose RefersTo points to a non‑existent sheet, deletes those named ranges, and saves the file. | Show how to safely parse a defined name's RefersTo string, handle quoted sheet names, and remove invalid named ranges after worksheets have been deleted in a .NET application.
// Common Searches: C# Aspose.Cells how to clean up defined names that point to deleted sheets | programmatically remove broken named ranges from Excel file using Aspose.Cells | detect and delete invalid named ranges after worksheet removal in .NET | Aspose.Cells iterate Names collection backwards to avoid index errors | remove named ranges referencing missing worksheets without throwing exceptions
// Tags: Aspose.Cells remove invalid named ranges | C# delete defined names with missing worksheet | iterate workbook.Names collection reverse | parse RefersTo sheet reference Aspose.Cells | clean Excel workbook broken named ranges .NET

using Aspose.Cells;

// The program loads an Excel workbook, walks the Names collection in reverse, extracts each defined name's RefersTo sheet reference (handling quoted names), checks if the worksheet exists, removes the named range when the sheet is absent, and saves the cleaned workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of all defined names (named ranges)
        var definedNames = workbook.Worksheets.Names;

        // Iterate backwards so removal does not affect the loop index
        for (int i = definedNames.Count - 1; i >= 0; i--)
        {
            Name name = definedNames[i];
            string refersTo = name.RefersTo;

            // Skip if the reference string is empty
            if (string.IsNullOrEmpty(refersTo))
                continue;

            // Remove leading '=' if present
            if (refersTo.StartsWith("="))
                refersTo = refersTo.Substring(1);

            // Find the position of '!' which separates sheet name from cell address
            int exclPos = refersTo.IndexOf('!');
            if (exclPos <= 0)
                continue; // Not a standard sheet reference

            // Extract the sheet name part
            string sheetName = refersTo.Substring(0, exclPos);

            // Remove surrounding single quotes (used when sheet name contains spaces)
            if (sheetName.StartsWith("'") && sheetName.EndsWith("'"))
                sheetName = sheetName.Substring(1, sheetName.Length - 2);

            // If the worksheet does not exist, remove the named range
            if (workbook.Worksheets[sheetName] == null)
                definedNames.RemoveAt(i);
        }

        // Save the cleaned workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
