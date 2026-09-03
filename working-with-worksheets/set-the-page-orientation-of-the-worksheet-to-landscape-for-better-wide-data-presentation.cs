// Title: Set a worksheet's page orientation to landscape with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that changes the PageSetup.Orientation of a worksheet to Landscape before saving the workbook using Aspose.Cells. | Generate a minimal Aspose.Cells example that creates a workbook, sets the first sheet to landscape orientation, and saves it as an .xlsx file.
// Common Searches: Aspose.Cells C# how to change worksheet orientation to landscape | set Excel sheet page orientation to landscape using Aspose.Cells .NET | C# example for PageSetup orientation landscape in Aspose.Cells | save workbook with landscape page layout Aspose.Cells API | modify page setup orientation of first worksheet Aspose.Cells C#
// Tags: Aspose.Cells configure worksheet orientation | C# landscape PageSetup Aspose.Cells | Excel workbook save with landscape orientation .NET | PageSetup orientation property Aspose.Cells | worksheet landscape mode C# Aspose

using Aspose.Cells;

// Creates a new workbook, accesses the first worksheet, sets its PageSetup.Orientation to Landscape, and saves the file as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the page orientation of the worksheet to landscape
        worksheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
