// Title: Read the paper width of the first worksheet in an XLSX workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get paper width of first worksheet from XLSX | how to read worksheet page setup paper width using Aspose.Cells | retrieve print layout dimensions of a worksheet with Aspose.Cells .NET | Aspose.Cells example to obtain worksheet paper size in points | load XLSX and query PageSetup.PaperWidth property in C#
// Tags: Aspose.Cells worksheet page setup paper width | C# read worksheet print dimensions from XLSX | Aspose.Cells get PageSetup.PaperWidth property | retrieve worksheet paper size using Aspose.Cells .NET | XLSX workbook paper width extraction with Aspose

using Aspose.Cells;

// The sample loads an XLSX workbook with Aspose.Cells, accesses the first worksheet, reads its PageSetup.PaperWidth property (in points), and prints the value to the console.
class Program
{
    static void Main()
    {
        // Load the XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet firstSheet = workbook.Worksheets[0];

        // Obtain the paper width (in points) of the first worksheet
        double paperWidth = firstSheet.PageSetup.PaperWidth;

        // Output the paper width
        System.Console.WriteLine($"Paper width of the first worksheet: {paperWidth} points");
    }
}
