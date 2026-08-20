// Title: C# – Create a Macro‑Free Workbook, Hide Zero Values, Set Top Margin, and Export to CSV with Aspose.Cells
// Description: Shows how to instantiate a macro‑free Aspose.Cells Workbook in C#, disable zero display, set the top page margin to 0.5 in, and save the worksheet directly as a CSV file.
// Keywords: Aspose.Cells C# hide zero values | Aspose.Cells set top margin | Aspose.Cells save as CSV | remove macros Aspose.Cells | macro free workbook Aspose.Cells | page setup Aspose.Cells | CSV export Aspose.Cells | C# Excel to CSV without macros
// Common Searches: Aspose.Cells hide zeros in worksheet | How to set page margins with Aspose.Cells .NET | Export Excel to CSV using Aspose.Cells C# | Remove macros from workbook Aspose.Cells | Create macro‑free workbook Aspose.Cells
// Developer Intent: Create a macro‑free Excel workbook, suppress zero values, apply a 0.5‑inch top margin, and generate a CSV file using Aspose.Cells for .NET.
// Use Cases: Generate CSV reports that omit zero entries for cleaner data analysis. | Prepare a printable worksheet with a specific top margin while still needing a CSV data dump. | Sanitize an existing Excel file by stripping macros before sharing the CSV version with external stakeholders. | Automate batch conversion of macro‑enabled workbooks to macro‑free CSV files in a CI pipeline.
// AI Prompts: Show C# code using Aspose.Cells to create a new workbook without macros, hide zero values, set the top margin to 0.5 inches, and save as CSV. | Give an example that removes macros, disables zero display, configures page setup margins, and exports the worksheet to CSV with Aspose.Cells for .NET.

using Aspose.Cells;

// Shows how to instantiate a macro‑free Aspose.Cells Workbook in C#, disable zero display, set the top page margin to 0.5 in, and save the worksheet directly as a CSV file.
class Program
{
    static void Main()
    {
        // Create a new workbook (macro‑free by default)
        Workbook workbook = new Workbook();

        // Remove any macros if present
        workbook.RemoveMacro();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Hide zero values in the worksheet
        sheet.DisplayZeros = false;

        // Set a custom top margin (in inches)
        sheet.PageSetup.TopMargin = 0.5; // 0.5 inch top margin

        // Save the workbook as a CSV file
        workbook.Save("output.csv", SaveFormat.Csv);
    }
}
