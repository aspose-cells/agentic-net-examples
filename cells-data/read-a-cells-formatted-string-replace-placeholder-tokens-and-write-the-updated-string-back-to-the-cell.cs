// Title: Replace {Name} and {Place} placeholders in an Excel cell using Aspose.Cells for .NET (C#)
// AI Prompts: Invoke Aspose.Cells Replace method with a custom ReplaceOptions object to substitute the {Name} token in cell A1 of a C# workbook. | Configure ReplaceOptions for case‑insensitive, partial‑cell matching and apply it to replace the {Place} placeholder alongside other tokens in the same cell. | After updating the cell's formatted string, call the Aspose.Cells Save method to write the workbook to an .xlsx file.
// Common Searches: aspnet replace placeholder text in Excel cell using Aspose.Cells C# | case insensitive string replacement in an Aspose.Cells worksheet cell | Aspose.Cells ReplaceOptions example for updating cell content | C# replace multiple tokens in a single Excel cell with Aspose.Cells | save workbook after modifying cell value with Aspose.Cells .NET
// Tags: replace placeholders in Excel cell Aspose.Cells C# | case insensitive ReplaceOptions Aspose.Cells | update formatted string cell Aspose.Cells | save workbook as Xlsx Aspose.Cells .NET | multiple token substitution worksheet cell

using System;
using Aspose.Cells;

namespace AsposeCellsPlaceholderReplacement
{
    // The example creates a workbook, writes a string containing {Name} and {Place} placeholders into cell A1, uses Aspose.Cells Replace with case‑insensitive ReplaceOptions to substitute each token, prints the updated value, and saves the workbook as PlaceholderReplaced.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put a sample string containing placeholders into cell A1
            // Example: "Hello {Name}, welcome to {Place}!"
            cells["A1"].PutValue("Hello {Name}, welcome to {Place}!");

            // Define replace options (case‑insensitive, replace within the cell content)
            ReplaceOptions options = new ReplaceOptions
            {
                CaseSensitive = false,
                MatchEntireCellContents = false
            };

            // Replace the placeholders one by one
            cells["A1"].Replace("{Name}", "Alice", options);
            cells["A1"].Replace("{Place}", "Wonderland", options);

            // Optionally, display the updated formatted string in the console
            Console.WriteLine("Updated cell value: " + cells["A1"].StringValue);

            // Save the workbook to a file
            workbook.Save("PlaceholderReplaced.xlsx", SaveFormat.Xlsx);
        }
    }
}
