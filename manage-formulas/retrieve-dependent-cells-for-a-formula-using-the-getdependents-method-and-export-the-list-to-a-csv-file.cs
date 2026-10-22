// Title: Extract dependent cells of a formula and save them to CSV using Aspose.Cells GetDependents in C#
// Common Searches: C# Aspose.Cells get dependent cells of a formula and export to CSV | How to list cells that depend on a specific formula with Aspose.Cells | Using GetDependents to retrieve formula dependencies in .NET | Save Aspose.Cells formula dependents to a CSV file in C# | Export worksheet cell dependency list to CSV using Aspose.Cells API
// Tags: Aspose.Cells GetDependents API | formula dependency extraction to CSV | C# Excel cell dependents enumeration | write dependent cell addresses to CSV | worksheet GetDependents usage example

using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// The sample loads an Excel workbook, selects a worksheet, and identifies a formula cell (default A1). It calls worksheet.Cells.GetDependents(false, row, column) to obtain all cells that reference that formula, builds a CSV string with the sheet name and each dependent cell address, and writes the output to 'dependents.csv', handling missing file and write errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "dependents.csv";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            var worksheet = workbook.Worksheets[0];

            // Specify the cell that contains the formula whose dependents you want
            var formulaCell = worksheet.Cells["A1"]; // change address as required

            // Retrieve dependent cells using GetDependents(includeHidden, row, column)
            Cell[] dependents = worksheet.Cells.GetDependents(false, formulaCell.Row, formulaCell.Column);

            // Build CSV content
            var sb = new StringBuilder();
            sb.AppendLine("SheetName,CellAddress");

            if (dependents != null)
            {
                foreach (var cell in dependents)
                {
                    string address = CellsHelper.CellIndexToName(cell.Row, cell.Column);
                    sb.AppendLine($"{worksheet.Name},{address}");
                }
            }

            // Write the CSV file
            try
            {
                File.WriteAllText(outputPath, sb.ToString());
                Console.WriteLine($"Dependent cells have been exported to {outputPath}");
            }
            catch (Exception writeEx)
            {
                Console.WriteLine($"Failed to write output file: {writeEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
