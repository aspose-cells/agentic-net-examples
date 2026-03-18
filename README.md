# Aspose.Cells for .NET – Agentic Code Examples

This repository contains automatically generated **Aspose.Cells for .NET** code examples.

These examples are produced by the **Aspose.Cells Product Agent** as part of the **Professionalize Multi-Agent Network for Bulk Code Example Generation**.

Only examples that **successfully build and run** are published.

## Architecture

The generation system consists of multiple agents:

| Agent | Port | Responsibility |
|------|------|------|
| Task Browser Agent | 7001 | Provides categories and tasks |
| Examples Super Agent | 7050 | Generates code examples |
| Aspose.Cells Product Agent | Local | Builds, tests and publishes examples |

## Example Generation Pipeline

The Aspose.Cells Product Agent performs the following steps:

1. Fetch categories from the **Task Browser Agent**
2. Fetch tasks for each category
3. Send tasks to the **Examples Super Agent**
4. Generate code examples
5. Build the example using `.NET`
6. Execute the example
7. Automatically fix errors using an LLM if needed
8. Publish validated examples to GitHub

Only **build-successful examples** are published.

## Repository Structure

Examples are organized by **category folders**.

agentic-net-examples

├ LICENSE  
├ README.md  
├ agents.md  
├ index.json  

├ worksheets  
│   ├ create-a-workbook.cs  
│   └ agents.md  

├ conversion  
│   ├ convert-xlsx-to-pdf.cs  
│   └ agents.md  

Each category folder contains:

- C# example files
- `agents.md` describing examples and usage patterns

## Build and Run

Each example is a simple **.NET console program**.

Build:

dotnet build

Run:

dotnet run

## Example File Naming Convention

All example files follow:

lowercase-with-hyphens.cs

Examples:

create-a-workbook.cs  
convert-xlsx-to-pdf.cs  
insert-column-chart.cs  

## Generated Files

This repository automatically generates the following files:

| File | Purpose |
|-----|-----|
| README.md | Repository overview |
| agents.md | Instructions for AI coding agents |
| index.json | Machine-readable index of examples |
| category/agents.md | Category-specific guidance |


## Contribution Workflow

Examples are generated automatically by the **Aspose.Cells Product Agent**.

The agent:

1. Pushes examples to a category branch
2. Creates a pull request
3. Maintainers review and merge the PR

Part of the **Professionalize Agent Network**.
