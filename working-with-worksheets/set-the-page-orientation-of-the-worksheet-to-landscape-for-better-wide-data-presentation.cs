// Title: How to set an Excel worksheet's page orientation to landscape using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to change the first worksheet's PageSetup orientation to Landscape and then saves the workbook. | Show the steps to configure a worksheet's print orientation to landscape in a .NET application with Aspose.Cells. | Generate a minimal Aspose.Cells example that sets a worksheet's page layout to landscape before exporting to an XLSX file.
// Common Searches: asp.net set worksheet page orientation to landscape with Aspose.Cells | c# Aspose.Cells change page setup orientation to landscape for existing workbook | how to print an Excel sheet in landscape mode using Aspose.Cells .NET | programmatically set worksheet print orientation to landscape in C# Aspose.Cells
// Tags: Aspose.Cells page orientation landscape | C# worksheet PageSetup orientation Aspose.Cells | Excel workbook landscape printing .NET | Aspose.Cells set worksheet print layout

using Aspose.Cells;

// Creates a new workbook (or loads one), accesses the first worksheet, sets its PageSetup.Orientation to Landscape, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // {CreateWorkbook}

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the page orientation to landscape for better wide data presentation
        sheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Save the workbook
        workbook.Save("Result.xlsx"); // {SaveWorkbook}
    }
}
