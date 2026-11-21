// Title: Set Excel worksheet zoom to 150% with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet zoom to 150 percent | programmatically change Excel sheet zoom level using Aspose.Cells | how to increase worksheet view scale to 150% in .NET | adjust zoom factor of a workbook worksheet with Aspose.Cells API | C# code to modify Excel worksheet zoom setting before saving
// Tags: worksheet zoom property Aspose.Cells | set worksheet view scale C# | adjust Excel sheet zoom programmatically | Aspose.Cells zoom level example | change workbook worksheet zoom .NET

using Aspose.Cells;

// Loads an existing workbook, accesses the first worksheet, sets its Zoom property to 150 (150 %), and saves the workbook to a new file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet you want to modify (e.g., the first worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the zoom factor to 150 percent for detailed visual inspection
        worksheet.Zoom = 150;

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
