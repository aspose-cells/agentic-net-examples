// Title: Insert a SUM formula in cell G10 to calculate the total of B2:B9 with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, sets cell G10 formula to =SUM(B2:B9) using Aspose.Cells, and saves as output.xlsx. | Show how to programmatically assign a SUM range formula to a specific cell (G10) in an Aspose.Cells worksheet.
// Common Searches: Aspose.Cells C# set formula =SUM(B2:B9) in G10 | How to add a SUM formula to a cell using Aspose.Cells .NET | Programmatically write Excel formula to cell G10 with Aspose.Cells | C# Aspose.Cells example inserting SUM formula into worksheet
// Tags: Aspose.Cells insert formula C# | set SUM formula Excel Aspose.Cells | write cell formula programmatically .NET | calculate range sum Aspose.Cells workbook | assign formula to specific cell Aspose.Cells

using Aspose.Cells;

// Creates a new workbook, accesses the first worksheet, assigns the formula =SUM(B2:B9) to cell G10, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert the SUM formula into cell G10 (calculates sum of B2:B9)
        Cell targetCell = worksheet.Cells["G10"];
        targetCell.Formula = "=SUM(B2:B9)";

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
