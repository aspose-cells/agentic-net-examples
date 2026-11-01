// Title: Ensuring Excel formula strings start with '=' before assigning them with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells add missing equal sign to formula before setting a cell | Why does Aspose.Cells throw a parsing error when a formula lacks the leading = sign | Programmatically validate that an Excel formula string begins with '=' using Aspose.Cells | How to automatically prepend '=' to formulas when using Aspose.Cells .NET | Set cell formula in Aspose.Cells without encountering an invalid formula error
// Tags: prepend equal sign to Excel formula Aspose.Cells | validate formula starts with equals Aspose.Cells C# | handle missing leading equals in Aspose.Cells formula assignment | Excel formula parsing error Aspose.Cells | set cell formula programmatically Aspose.Cells .NET

using Aspose.Cells;

// The example creates a workbook, checks a formula string for a leading '=', adds it if missing, assigns the corrected formula to cell B1, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Access the first worksheet
        var sheet = workbook.Worksheets[0];

        // Define the formula string (without leading equal sign)
        string formula = "SUM(A1:A10)";

        // Ensure the formula starts with an equal sign
        if (!formula.StartsWith("="))
        {
            formula = "=" + formula;
        }

        // Set the formula to cell B1
        sheet.Cells["B1"].Formula = formula;

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
