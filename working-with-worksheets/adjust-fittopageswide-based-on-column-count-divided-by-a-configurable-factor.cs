// Title: Set worksheet FitToPagesWide dynamically from column count using a configurable divisor with Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads the used column range of a worksheet, divides it by a user‑defined factor, rounds up, and assigns the result to PageSetup.FitToPagesWide while keeping FitToPagesTall set to zero. | Show how to expose the page‑width scaling factor as a runtime parameter when printing an Excel file with Aspose.Cells. | Create a reusable method that takes a Workbook and a divisor, computes the required pages wide based on MaxColumn, updates the worksheet's page setup, and saves the file.
// Common Searches: how to calculate FitToPagesWide from column count in Aspose.Cells C# | Aspose.Cells set print scaling based on number of columns and custom factor | C# dynamic page width for Excel printing using divisor Aspose.Cells | adjust worksheet FitToPagesWide automatically according to used columns .NET
// Tags: Aspose.Cells FitToPagesWide calculation | worksheet page setup scaling factor | C# column count based print layout | dynamic page width divisor Aspose.Cells | Excel printing FitToPagesTall zero .NET

using Aspose.Cells;
using System;

// Load the workbook (replace with your actual file path)
var workbook = new Workbook("input.xlsx");

// Configurable factor to divide the column count
double factor = 2.0; // Adjust this value as needed

// Access the first worksheet (or specify the desired index)
var worksheet = workbook.Worksheets[0];

// Determine the number of columns that contain data
int columnCount = worksheet.Cells.MaxColumn + 1; // MaxColumn is zero‑based

// Calculate the number of pages wide, rounding up to ensure all columns fit
int pagesWide = (int)Math.Ceiling(columnCount / factor);

// Apply the FitToPagesWide setting; set FitToPagesTall to 0 to let height adjust automatically
worksheet.PageSetup.FitToPagesWide = pagesWide;
worksheet.PageSetup.FitToPagesTall = 0;

// Save the modified workbook (replace with your desired output path)
workbook.Save("output.xlsx");
