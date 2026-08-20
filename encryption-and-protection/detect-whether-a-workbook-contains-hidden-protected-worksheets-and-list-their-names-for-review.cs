// Title: Find Hidden and Protected Worksheets in an Excel Workbook with Aspose.Cells for .NET (C#)
// Description: Loads an Excel file using Aspose.Cells, iterates through every worksheet, checks hidden status via IsVisible/VisibilityType and protection via IsProtected, and prints the names of sheets that are both hidden and protected.
// Keywords: Aspose.Cells hidden worksheets | Aspose.Cells protected sheets | C# detect hidden Excel sheets | list hidden protected worksheets .NET | Excel workbook security audit Aspose | enumerate hidden protected worksheets C# | worksheet visibility Aspose.Cells | check worksheet protection Aspose.Cells
// Common Searches: how to list hidden and password protected worksheets using Aspose.Cells C# | Aspose.Cells C# find worksheets that are hidden and protected | detect hidden protected sheets in Excel with Aspose.Cells .NET | enumerate invisible protected worksheets Aspose.Cells | C# code to check worksheet visibility and protection Aspose
// Developer Intent: Retrieve the names of all worksheets that are simultaneously hidden and protected in an Excel workbook.
// Use Cases: Perform a security audit of an Excel file before distribution by identifying concealed protected sheets. | Automate compliance checks to ensure confidential data isn’t stored in hidden protected worksheets. | Generate a report of hidden protected worksheets for documentation or review purposes.
// AI Prompts: Write C# code with Aspose.Cells that lists hidden and protected worksheets and then unhides them. | Show how to save a copy of the workbook after removing protection from hidden sheets using Aspose.Cells. | Explain the difference between hidden, very hidden, and protected worksheets and the Aspose.Cells properties that expose each state.

using System;
using Aspose.Cells;

namespace AsposeCellsHiddenProtectedSheets
{
    // Loads an Excel file using Aspose.Cells, iterates through every worksheet, checks hidden status via IsVisible/VisibilityType and protection via IsProtected, and prints the names of sheets that are both hidden and protected.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be inspected
            string inputPath = "input.xlsx";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            Console.WriteLine("Hidden and protected worksheets:");

            // Iterate through all worksheets
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                // Determine if the sheet is hidden (IsVisible false or VisibilityType not Visible)
                bool isHidden = !sheet.IsVisible || sheet.VisibilityType != VisibilityType.Visible;

                // Determine if the sheet is protected
                bool isProtected = sheet.IsProtected;

                // If both conditions are true, output the sheet name
                if (isHidden && isProtected)
                {
                    Console.WriteLine($"- {sheet.Name}");
                }
            }

            // Optionally, save a copy of the workbook (lifecycle rule: save)
            // workbook.Save("output.xlsx");
        }
    }
}
