// Title: Read the paper width of the first worksheet from an XLSX workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an XLSX file with Aspose.Cells and outputs the PageSetup.PaperWidth of the first worksheet. | Show how to access the worksheet's PageSetup property to retrieve its printing paper width in inches using Aspose.Cells. | Adapt the example to convert the obtained paper width from inches to centimeters and display the result.
// Common Searches: Aspose.Cells C# get paper width of first worksheet in Excel file | How to read worksheet page setup dimensions with Aspose.Cells .NET | Retrieve printing paper width property from an Excel sheet using Aspose.Cells | C# example for Aspose.Cells PageSetup PaperWidth
// Tags: aspacells page-setup paperwidth property | c# aspacells workbook worksheet access | first worksheet printing dimensions aspacells | inch to centimeter conversion c# utility | aspacells example reading page setup values

using Aspose.Cells;

// Loads an XLSX workbook with Aspose.Cells, accesses the first worksheet, reads its PageSetup.PaperWidth (in inches), and prints the value.
class Program
{
    static void Main()
    {
        // Load the XLSX workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Obtain the paper width (in inches) from the worksheet's page setup
        double paperWidth = worksheet.PageSetup.PaperWidth;

        // Output the paper width
        System.Console.WriteLine($"Paper width of the first worksheet: {paperWidth} inches");
    }
}
