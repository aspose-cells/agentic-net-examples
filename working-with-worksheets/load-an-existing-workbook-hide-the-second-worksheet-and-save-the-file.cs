// Title: Hide the second worksheet in an existing Excel workbook with Aspose.Cells for .NET and save the changes
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets the second worksheet’s visibility to hidden, and saves the workbook to a new file. | Show how to use Aspose.Cells to programmatically hide a worksheet by index and persist the modified workbook in .NET. | Provide a snippet that demonstrates changing a worksheet’s IsVisible property to false and then exporting the workbook using Aspose.Cells.
// Common Searches: aspnet hide worksheet at index 1 using Aspose.Cells | C# Aspose.Cells hide second sheet and save workbook | how to make a specific worksheet invisible in an existing Excel file with Aspose.Cells .NET | Aspose.Cells hide worksheet programmatically before saving
// Tags: hide worksheet by index Aspose.Cells .NET | set worksheet IsVisible false Aspose.Cells | save modified workbook after changing sheet visibility | Aspose.Cells hide second sheet example | programmatic worksheet visibility control Aspose.Cells

using Aspose.Cells;

// Loads 'input.xlsx' with Aspose.Cells, hides the worksheet at index 1, and saves the result as 'output.xlsx'.
class Program
{
    static void Main()
    {
        // Load the existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the second worksheet (index 1) and hide it
        Worksheet secondSheet = workbook.Worksheets[1];
        secondSheet.IsVisible = false; // Alternatively: secondSheet.Visibility = VisibilityType.Hidden;

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
