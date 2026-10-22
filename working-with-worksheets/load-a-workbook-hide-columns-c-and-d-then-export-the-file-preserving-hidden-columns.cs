// Title: Hide columns C and D in an Excel workbook and preserve their hidden state when saving using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing .xlsx file with Aspose.Cells, hide column C (index 2) and column D (index 3) in the first worksheet, then save the workbook so the hidden columns remain hidden. | Using Aspose.Cells for .NET, programmatically set column visibility to hidden for specific columns and export the workbook without revealing those columns.
// Common Searches: C# Aspose.Cells hide specific columns and keep them hidden after saving the workbook | How to preserve hidden column settings when exporting Excel with Aspose.Cells .NET | Aspose.Cells hide column by index example C# | Save Excel file with hidden columns using Aspose.Cells for .NET
// Tags: hide column by index Aspose.Cells | preserve column visibility on workbook save | Aspose.Cells hide columns C D | export Excel with hidden columns C# | column visibility control Aspose.Cells .NET

using Aspose.Cells;

// Loads 'input.xlsx', hides columns C (index 2) and D (index 3) on the first worksheet, and saves as 'output.xlsx' while retaining the hidden column settings.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (you can change the index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide column C (zero‑based index 2) and column D (index 3)
        worksheet.Cells.HideColumn(2);
        worksheet.Cells.HideColumn(3);

        // Save the workbook; hidden columns are preserved in the saved file
        workbook.Save("output.xlsx");
    }
}
