// Title: Set FitToPagesWide and FitToPagesTall to zero for unlimited page scaling with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set FitToPagesWide to 0 for unlimited width | How to remove page height limit using FitToPagesTall in Aspose.Cells | Print scaling without page count limit Aspose.Cells .NET example | Configure worksheet page setup to ignore page limits in C# | Aspose.Cells unlimited page dimensions for Excel export
// Tags: FitToPagesWide zero unlimited scaling | FitToPagesTall zero unlimited scaling | Aspose.Cells page setup print scaling | C# Aspose.Cells worksheet page configuration | unlimited page dimensions Aspose.Cells

using Aspose.Cells;

// The example creates a workbook, accesses the first worksheet's PageSetup, sets FitToPagesWide and FitToPagesTall to 0 to allow unlimited page width and height, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the page setup object for the worksheet
        PageSetup pageSetup = worksheet.PageSetup;

        // Set FitToPagesWide and FitToPagesTall to 0 for unlimited pages
        pageSetup.FitToPagesWide = 0;   // Unlimited page width
        pageSetup.FitToPagesTall = 0;   // Unlimited page height

        // Save the workbook to a file
        workbook.Save("Result.xlsx");
    }
}
