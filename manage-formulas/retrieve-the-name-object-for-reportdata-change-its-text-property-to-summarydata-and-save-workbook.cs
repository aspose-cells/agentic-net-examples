// Title: Rename a Named Range (ReportData → SummaryData) in Aspose.Cells for .NET and Save the Workbook
// Description: Shows how to get a Name object, update its Text property from "ReportData" to "SummaryData", and save the workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells rename named range | Name.Text property C# | modify named range Aspose.Cells | Workbook.Save after rename | C# Aspose.Cells named range management
// Common Searches: rename named range Aspose.Cells .NET | change Name.Text Aspose.Cells | update named range and save workbook C# | Aspose.Cells how to rename a name object | C# code to modify named range in Excel file
// Developer Intent: Change the existing named range "ReportData" to "SummaryData" and write the updated workbook to disk.
// Use Cases: Align workbook named ranges with a standardized reporting schema before distribution. | Allow end‑users to rename ranges dynamically based on UI input during runtime. | Prepare a batch of workbooks for automated processing by enforcing a consistent naming convention.
// AI Prompts: Write C# code using Aspose.Cells that checks if a named range exists, renames it safely, and saves the workbook, handling errors gracefully. | Create a reusable method that accepts oldName, newName, and filePath, renames the Name object, and persists the workbook. | Explain the difference between Name.Text and Name.RefersTo in Aspose.Cells, with examples of when to use each.

using Aspose.Cells;

// Shows how to get a Name object, update its Text property from "ReportData" to "SummaryData", and save the workbook with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Ensure there is a named range called "ReportData"
        // (If it already exists, this will add another with the same name; adjust as needed)
        int idx = workbook.Worksheets.Names.Add("ReportData");
        Name reportName = workbook.Worksheets.Names[idx];
        reportName.RefersTo = "=Sheet1!$A$1:$B$2";

        // Retrieve the Name object for "ReportData"
        Name nameObj = workbook.Worksheets.Names["ReportData"];

        // Change its Text property to "SummaryData"
        nameObj.Text = "SummaryData";

        // Save the workbook
        workbook.Save("ModifiedWorkbook.xlsx");
    }
}
