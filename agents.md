# Aspose.Cells Product Agent

This repository is maintained automatically by the **Aspose.Cells Product Agent**.

The agent is part of the **Professionalize Multi-Agent Network for Bulk Code Example Generation**.

## Agent Responsibilities

The Aspose.Cells Product Agent performs the following tasks:

1. Fetches categories and tasks from the **Task Browser Agent**
2. Requests example code generation from the **Examples Super Agent**
3. Builds and validates generated code examples
4. Applies automated fixes using an LLM when needed
5. Publishes validated examples to the appropriate category folders
6. Creates pull requests for review

Only examples that successfully **build and run** are published.

## Repository Organization

Examples are grouped into folders by category.

Each category contains:

- C# example files
- `agents.md` listing the tasks implemented in that category

Example categories include:

- Conversion
- Worksheets
- Charts
- Formatting
- PivotTables

## Review Workflow

The Aspose.Cells Product Agent pushes examples using a dedicated agent account and creates pull requests.

Repository maintainers review and merge the pull requests.

---

Part of the **Professionalize Agent Network**.
