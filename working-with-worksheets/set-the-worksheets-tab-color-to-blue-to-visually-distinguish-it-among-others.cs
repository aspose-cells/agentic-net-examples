// Title: Set worksheet tab color to blue in C# with Aspose.Cells
// Description: This C# example demonstrates creating a workbook, selecting the first worksheet, assigning its TabColor property to blue, and saving the result as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# tab color | set worksheet tab color | Excel sheet tab color blue | Aspose.Cells TabColor property | change Excel worksheet tab color programmatically | C# Aspose.Cells workbook styling
// Common Searches: C# Aspose.Cells set worksheet tab color | how to change Excel sheet tab color with Aspose.Cells | set tab color to blue in generated workbook C# | Aspose.Cells change worksheet tab color example | programmatically color Excel worksheet tabs .NET
// Developer Intent: Apply a blue color to a worksheet’s tab in an Excel file using Aspose.Cells for .NET.
// Use Cases: Highlight a summary sheet with a blue tab for quick visual identification. | Assign distinct colors to data sheets in an automated report to improve navigation. | Generate workbooks with colored tabs to enhance end‑user experience in Excel.
// AI Prompts: Show me how to set the tab color of a specific worksheet to a custom RGB value using Aspose.Cells for .NET. | Provide code to change the tab color of multiple worksheets based on their names with Aspose.Cells. | Explain how to retrieve and modify the existing tab color of a worksheet in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# example demonstrates creating a workbook, selecting the first worksheet, assigning its TabColor property to blue, and saving the result as an XLSX file using Aspose.Cells for .NET.
    public class SetWorksheetTabColor
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set the worksheet tab color to blue
                worksheet.TabColor = Color.Blue;

                // Save the workbook
                string outputPath = "WorksheetWithBlueTab.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            SetWorksheetTabColor.Run();
        }
    }
}
