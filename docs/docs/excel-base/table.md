# Tables

Tables are one of the most powerful and versatile data structures in the XlBlocks framework. They allow you to work with structured, tabular data directly within Excel, providing a wide range of functionality for creating, manipulating, and analyzing data. Tables are represented in Excel using object handles, enabling seamless integration with other XlBlocks objects.

---

## Key Features

- **Building Tables**: Create tables from Excel ranges, delimited files (e.g., CSV), or other data sources.
- **Data Manipulation**: Perform operations such as filtering, sorting, joining, and grouping.
- **Column Operations**: Access, modify, or append columns, including support for derived columns using expressions and data from dictionaries or lists.
- **Exporting**: Save tables to delimited files for external use.
- **Integration**: Convert tables to dictionaries or lists, and vice versa, for flexible data workflows.

---

## Example Use Cases

1. **Data Cleaning**: Use filtering and null-dropping functions to clean raw data.
2. **Data Analysis**: Perform grouping and aggregation to calculate statistics.
3. **Data Transformation**: Append derived columns or join multiple tables for enriched datasets.
4. **Data Export**: Save processed tables to CSV for sharing or further analysis.

---

## Expression Language

The XlBlocks expression language is a powerful tool for filtering rows and adding new fields to tables. It allows you to write expressions that are evaluated row by row, enabling dynamic and flexible data manipulation. This is expression language is used by the functions [`XBTable_FilterWith`](#xbtable_filterwith) and [`XBTable_AppendColumnsWith`](#xbtable_filterwith). 

For more information, refer to the [Expression Language Reference](../expressions.md).

---

## Table Functions
