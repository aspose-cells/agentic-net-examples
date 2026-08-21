// Title: Protect an Excel worksheet with Aspose.Cells for .NET and export to CSV
// Description: Creates a workbook, fills cells A1:B3, applies full worksheet protection (no password) using ProtectionType.All, and saves the sheet as a CSV file. The protection setting does not modify the exported CSV content.
// Keywords: Aspose.Cells | worksheet protection | ProtectionType.All | export to CSV | SaveFormat.Csv | .NET Excel | protected worksheet export | no password protection | Excel to CSV conversion
// Common Searches: Aspose.Cells protect worksheet and export CSV | Does worksheet protection affect CSV output in Aspose.Cells | Save protected Excel sheet as CSV .NET | Export protected worksheet to CSV without password | How to keep cell values when saving protected sheet as CSV
// Developer Intent: Apply worksheet protection and generate a CSV file that retains all original cell values.
// Use Cases: Lock a sheet to prevent editing in Excel while still providing a CSV report for downstream systems. | Distribute a protected workbook to users but automate CSV extraction for data pipelines. | Secure worksheet layout without a password and produce unchanged CSV files for integration with third‑party tools.
// AI Prompts: Show C# code that protects an Excel worksheet with Aspose.Cells and then saves it as CSV. | Explain whether worksheet protection influences the CSV result when using SaveFormat.Csv in Aspose.Cells. | Provide an example of protecting a worksheet with a password and ensuring the CSV export includes all cell values.

using Aspose.Cells;

// Creates a workbook, fills cells A1:B3, applies full worksheet protection (no password) using ProtectionType.All, and saves the sheet as a CSV file. The protection setting does not modify the exported CSV content.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("Alice");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Bob");
        sheet.Cells["B3"].PutValue(25);

        // Protect the worksheet (all protection types, no password)
        sheet.Protect(ProtectionType.All);

        // Export the worksheet data to CSV.
        // The protection setting does not affect the exported values.
        workbook.Save("ProtectedWorksheet.csv", SaveFormat.Csv);
    }
}
