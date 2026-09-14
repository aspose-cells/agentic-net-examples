// Title: Set worksheet zoom to 150% using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to set the Zoom property of a worksheet to 150 and save the workbook. | Programmatically change the view scale of the first worksheet to 150% with Aspose.Cells for .NET. | Apply a 150 percent zoom factor to an Excel sheet using Aspose.Cells and output the file as output.xlsx.
// Common Searches: Aspose.Cells C# how to change worksheet view zoom to 150 percent | set Excel sheet zoom level programmatically with Aspose.Cells .NET | increase worksheet zoom factor to 150% before saving workbook in C# | adjust worksheet Zoom property using Aspose.Cells example code | C# Aspose.Cells set zoom for first worksheet for detailed inspection
// Tags: worksheet zoom property Aspose.Cells | apply 150 percent zoom C# | adjust Excel view scale Aspose.Cells | programmatic worksheet zoom Aspose.Cells .NET | increase sheet visual inspection zoom

using Aspose.Cells;

// Creates (or loads) a workbook, accesses the first worksheet, sets its Zoom property to 150 for detailed visual inspection, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // create

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Adjust the zoom factor to 150 percent for detailed visual inspection
        sheet.Zoom = 150;

        // Save the workbook to a file
        workbook.Save("output.xlsx"); // save
    }
}
