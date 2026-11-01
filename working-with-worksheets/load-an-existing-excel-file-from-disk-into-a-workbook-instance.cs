// Title: Load an existing Excel workbook from a local .xlsx file using Aspose.Cells in C#
// AI Prompts: Write C# code that creates an Aspose.Cells Workbook by providing the full path to a .xlsx file on disk. | Show how to open a local Excel file with Aspose.Cells and obtain a Workbook object for further processing. | Demonstrate loading an existing spreadsheet into a Workbook instance using Aspose.Cells in a .NET application.
// Common Searches: aspocells c# load workbook from local file path | how to open a .xlsx file with Aspose.Cells in a .NET project | C# read existing Excel workbook using Aspose.Cells library | instantiate Aspose.Cells Workbook with absolute file location | load Excel data into Aspose.Cells object from disk
// Tags: Aspose.Cells load workbook from .xlsx file | C# instantiate Workbook with file path | Aspose.Cells open existing Excel workbook | read local Excel file using Aspose.Cells | load spreadsheet into Aspose.Cells object

using Aspose.Cells;

// Specify the path to the existing Excel file
string filePath = @"C:\Data\Sample.xlsx";

// Load the workbook from the file
Workbook workbook = new Workbook(filePath);

// At this point, 'workbook' contains the loaded Excel data and can be manipulated as needed.
