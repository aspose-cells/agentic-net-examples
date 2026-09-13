// Title: Enable black‑and‑white printing for every worksheet in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file, iterates through all worksheets, sets PageSetup.BlackAndWhite to true, and saves the workbook. | Show how to apply grayscale printing mode to each sheet in a workbook using the Aspose.Cells API. | Provide a loop example that configures black‑and‑white page setup for every worksheet in a workbook and writes the output file.
// Common Searches: Aspose.Cells C# set black and white printing for all worksheets in a workbook | how to configure grayscale page setup for every sheet using Aspose.Cells | C# code to reduce ink usage by enabling black‑and‑white printing in Excel files | apply black and white printing mode to entire workbook with Aspose.Cells .NET | bulk update worksheet page setup to black and white using Aspose.Cells library
// Tags: set worksheet black and white page setup Aspose.Cells | apply grayscale printing to workbook C# | bulk page setup configuration Aspose.Cells | reduce ink consumption Excel Aspose.Cells | activate black and white printing across workbook | Aspose.Cells PageSetup.BlackAndWhite property

using Aspose.Cells;

// Load the workbook (replace with your actual file path)
Workbook workbook = new Workbook("input.xlsx");

// Enable black‑and‑white printing for every worksheet in the workbook
foreach (Worksheet sheet in workbook.Worksheets)
{
    sheet.PageSetup.BlackAndWhite = true;
}

// Save the workbook (replace with your desired output path)
workbook.Save("output.xlsx");
