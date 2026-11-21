// Title: How to add a calculated column that joins two text columns with a hyphen in an Excel file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add calculated column that combines column A and B with '-' | C# formula to concatenate two cells with a dash in an existing workbook using Aspose.Cells | Insert a hyphen‑separated full name column in Excel via Aspose.Cells API | Using Aspose.Cells to create a new column with =A2&"-"&B2 formula in .NET | Programmatically add a concatenated text field to a worksheet with Aspose.Cells for .NET
// Tags: add calculated column Aspose.Cells | concatenate text cells formula .NET | Aspose.Cells insert hyphen separator column | Excel workbook formula insertion C# | recalculate formulas after adding column Aspose.Cells

using Aspose.Cells;

// The example loads an existing workbook, adds a header "FullName" in column C, sets a formula that joins the values of columns A and B with a hyphen for each data row, recalculates all formulas, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load the existing workbook (create/load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Determine the last row that contains data
        int lastDataRow = sheet.Cells.MaxDataRow;

        // Add header for the calculated field in column C (index 2)
        sheet.Cells[0, 2].PutValue("FullName");

        // Set the concatenation formula for each data row
        for (int row = 1; row <= lastDataRow; row++)
        {
            // Formula: =A{row+1}&"-"&B{row+1}
            sheet.Cells[row, 2].Formula = $"=A{row + 1}&\"-\"&B{row + 1}";
        }

        // Recalculate all formulas so the new column is populated
        workbook.CalculateFormula();

        // Save the workbook with the new calculated field (save rule)
        workbook.Save("output.xlsx");
    }
}
