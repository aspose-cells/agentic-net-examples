// Title: Add a modify‑password to an existing ODS workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an ODS file, assigns a password required for modification via Workbook.Settings.Password, and saves the workbook unchanged with Aspose.Cells. | Explain how to protect an existing ODS spreadsheet from editing by applying a modify password in Aspose.Cells for .NET. | Provide a step‑by‑step example of using Aspose.Cells to set an edit‑only password on an ODS workbook while preserving all original data.
// Common Searches: Aspose.Cells C# set modify password on existing .ods workbook | protect ODS file from editing while keeping data using .NET | how to apply edit‑only password to an ODS spreadsheet with Aspose.Cells | save ODS workbook with password‑to‑modify without data loss in C#
// Tags: ODS modify password Aspose.Cells | Aspose.Cells password setting API | C# protect ODS workbook edit | Aspose.Cells save ODS with password | preserve data ODS encryption .NET

using Aspose.Cells;

// Load the existing ODS workbook
Workbook workbook = new Workbook("input.ods");

// Apply a password‑to‑modify protection
workbook.Settings.Password = "myPassword";

// Save the workbook, preserving all original data
workbook.Save("output.ods", SaveFormat.Ods);
