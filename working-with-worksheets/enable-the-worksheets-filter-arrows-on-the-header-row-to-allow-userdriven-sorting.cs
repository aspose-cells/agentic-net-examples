// Title: Enable AutoFilter arrows on the header row of an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new workbook, writes a header row, sets the AutoFilter range to the header cells, and saves the file with Aspose.Cells. | Write a C# snippet that determines the last used column in the first worksheet and applies AutoFilter arrows to that entire header row using Aspose.Cells. | Provide C# code to apply a specific filter condition (e.g., Age > 25) after enabling AutoFilter arrows on the header row with Aspose.Cells. | Show how to clear or remove the AutoFilter from a worksheet programmatically in C# using Aspose.Cells.
// Common Searches: aspnet how to add filter arrows to the first row of an Excel file with Aspose.Cells | C# Aspose.Cells enable AutoFilter on dynamic column range | example code for setting AutoFilter range A1:D1 using Aspose.Cells .NET | apply Excel AutoFilter programmatically with Aspose.Cells C# tutorial | Aspose.Cells filter arrows on header row without using Excel UI
// Tags: Aspose.Cells enable AutoFilter C# | C# set AutoFilter range Aspose.Cells | Aspose.Cells worksheet filter arrows | Excel AutoFilter programmatic Aspose.Cells | dynamic column AutoFilter Aspose.Cells .NET

using Aspose.Cells;

// The example creates a new workbook, adds a header row (Name, Age, Country, Score), inserts sample data, enables AutoFilter arrows on the header range A1:D1, and saves the workbook as FilteredWorkbook.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Access the first worksheet
        var sheet = workbook.Worksheets[0];

        // Populate header row
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["C1"].PutValue("Country");
        sheet.Cells["D1"].PutValue("Score");

        // Add sample data rows
        sheet.Cells["A2"].PutValue("Alice");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["C2"].PutValue("USA");
        sheet.Cells["D2"].PutValue(85);

        sheet.Cells["A3"].PutValue("Bob");
        sheet.Cells["B3"].PutValue(25);
        sheet.Cells["C3"].PutValue("UK");
        sheet.Cells["D3"].PutValue(90);

        // Enable filter arrows on the header row (A1:D1)
        sheet.AutoFilter.Range = "A1:D1";

        // Save the workbook
        workbook.Save("FilteredWorkbook.xlsx");
    }
}
