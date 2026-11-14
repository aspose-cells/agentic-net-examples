// Title: Insert a row‑specific total formula (Quantity × UnitPrice) into an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to add a formula that multiplies two columns for each row in Aspose.Cells C# | Aspose.Cells C# set cell formula dynamically based on row number | calculate total column from quantity and unit price using Aspose.Cells .NET | programmatically write =B3*C3 formulas to Excel with Aspose.Cells | loop through rows and assign formulas in Aspose.Cells workbook C#
// Tags: Aspose.Cells total column formula generation | C# row-wise formula assignment Excel | Excel cell formula creation Aspose.Cells | quantity unitprice multiplication Aspose.Cells | automated Excel formula insertion .NET

using Aspose.Cells;

// The example loads an existing workbook, identifies the Quantity (B), Unit Price (C), and Total (D) columns, iterates over each data row, builds a row‑specific formula like =B3*C3, assigns it to the Total cell, and saves the updated Excel file.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Define column indexes:
        // Quantity column -> B (index 1)
        // Unit Price column -> C (index 2)
        // Total column -> D (index 3)
        int quantityCol = 1;
        int unitPriceCol = 2;
        int totalCol = 3;

        // Determine the range of rows that contain data.
        // Assuming the first row (0) holds headers, start from row 1.
        int startRow = 1;
        int lastRow = sheet.Cells.MaxDataRow; // last row with data

        // Insert the formula for each data row.
        for (int row = startRow; row <= lastRow; row++)
        {
            // Build the formula string for the current row.
            // Example for row 2 (Excel row 3): =B3*C3
            string formula = $"=B{row + 1}*C{row + 1}";

            // Assign the formula to the Total cell.
            sheet.Cells[row, totalCol].Formula = formula;
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
