# Expression Language Reference

The XlBlocks expression language is a powerful tool for filtering rows and adding new fields to tables. It allows you to write expressions that are evaluated row by row, enabling dynamic and flexible data manipulation.

---

## Overview

The expression language supports:

- **Filtering**: Use expressions in functions like `XBTable_FilterWith` to filter rows based on conditions.
- **Field Creation**: Add new fields to tables by evaluating expressions for each row.

Expressions can reference columns, use literals, apply operators, and call functions.

---

## Syntax

### Column References

- Columns are referenced using square brackets: `[ColumnName]`.
- Example: `[Name]`, `[Age]`.

### Literals

- **String Literals**: Enclosed in single quotes: `'string'`.
- **Numeric Literals**: Integers (`42`) or floating-point numbers (`3.14`).
- **Boolean Literals**: `TRUE`, `FALSE`.
- **Null Literal**: `NULL` represents a missing value.

### Expressions

- Expressions can combine columns, literals, operators, and functions.
- Example: `[Age] + 10`, `IIF([Name] == 'Charlie', 'Yes', 'No')`.

---

## Operators

The operators defined in the table below are sorted in order of higher to lower precedence. Associativity determines the order of evaluation for operators with the same precedence.

| Operator                         | Description                      |
|----------------------------------|----------------------------------|
| `(`, `)`                         | Parentheses                      |
| `NOT`, `!`                       | Logical NOT                      |
| `^`                              | Exponentiation                   |
| `*`, `/`, `%`                    | Multiplication, Division, Modulo |
| `+`, `-`                         | Addition, Subtraction            |
| `==`, `!=`, `<`, `>`, `<=`, `>=` | Comparison Operators             |
| `IS`                             | Null Check                       |
| `AND`                            | Logical AND                      |
| `OR`, `XOR`                      | Logical OR, XOR                  |

---

## Functions

### General Syntax

Functions are called with arguments enclosed in parentheses: `FUNCTION(arg1, arg2, ...)`.

### Expression Functions

| Function        | Description                                                                                    | Usage ([bracketed] args are optional)                                    |
|-----------------|------------------------------------------------------------------------------------------------|--------------------------------------------------------------------------|
| `IIF`           | Ternary function                                                                               | `IIF(condition, trueExpression, falseExpression)`                        |
| `ISNULL`        | Returns a value if the column is null                                                          | `ISNULL(column, valueIfNull)`                                            |
| `LEN`           | Returns the length of a string                                                                 | `LEN(column)`                                                            |
| `EXP`           | Computes an exponent of a column                                                               | `EXP([base], exponent)` or `EXP(exponent)` (default base is `e`)         |
| `LOG`           | Computes the logarithm of a column                                                             | `LOG(value, [base])` or `LOG(value)` (default base is `e`)               |
| `CUMSUM`        | Cumulative sum, optionally conditioned on another column                                       | `CUMSUM(column, [condition])`                                            |
| `CUMPROD`       | Cumulative product, optionally conditioned on another column                                   | `CUMPROD(column, [condition])`                                           |
| `CUMMIN`        | Cumulative min, optionally conditioned on another column                                       | `CUMMIN(column, [condition])`                                            |
| `CUMMAX`        | Cumulative max, optionally conditioned on another column                                       | `CUMMAX(column, [condition])`                                            |
| `ROUND`         | Rounds a value to a number of digits, positive digits are after decimal, negative are before   | `ROUND(column, [digits])`                                                |
| `ABS`           | Extracts a substring                                                                           | `SUBSTRING(column, start, [length])`                                     |
| `MIN`           | Computes the minimum value between two columns                                                 | `MIN(column, otherColumn)`                                               |
| `MAX`           | Computes the maximum value between two columns                                                 | `MAX(column, otherColumn)`                                               |
| `FLOOR`         | Returns the nearest integer below the column value                                             | `FLOOR(column)`                                                          |
| `CEILING`       | Returns the nearest integer above the column value                                             | `CEILING(column)`                                                        |
| `SUBSTRING`     | Extracts a substring                                                                           | `SUBSTRING(column, start, [length])`                                     |
| `LEFT`          | Extracts a number of characters from the start of a string                                     | `LEFT(column, length)`                                                   |
| `RIGHT`         | Extracts a number of characters from the end of a string                                       | `RIGHT(column, length)`                                                  |
| `TRIM`          | Removes leading and trailing whitespace from a string                                          | `TRIM(column)`                                                           |
| `REPLACE`       | Replaces instances of a value in a string                                                      | `REPLACE(column, oldValue, newValue, [caseSensitive])`                   |
| `REGEX_TEST`    | Tests a string against a regex pattern &dagger;                                                | `REGEX_TEST(column, pattern, caseSensitive)`                             |
| `REGEX_FIND`    | Finds a value in a string using a regex pattern &dagger;                                       | `REGEX_FIND(column, pattern, caseSensitive)`                             |
| `REGEX_REPLACE` | Replaces a value in a string using a regex pattern &dagger;                                    | `REGEX_REPLACE(column, pattern, caseSensitive)`                          |
| `FORMAT`        | Converts a value to a string using the specified format &dagger;                               | `FORMAT(column, formatString)`                                           |
| `TODATE`        | Converts a value to a date &dagger;                                                            | `TODATE(column, formatString)`                                           |

> &dagger; As XlBlocks is built on .NET, the Regex and format strings for the `REGEX_*`, `FORMAT` and `TODATE` functions are based on .NET's implementations:
>
> - [Regex Pattern Reference](https://learn.microsoft.com/en-us/dotnet/standard/base-types/regular-expression-language-quick-reference)
> - [String Formatting Reference](https://learn.microsoft.com/en-us/dotnet/standard/base-types/standard-numeric-format-strings)
> - [Date and Time Format Strings](https://learn.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings)

---

## Special Constructs

### `IN`

- Checks if a value exists in a list, evaluating to a boolean.
- Syntax: `[ColumnName] IN ('value1', 'value2', ...)`
- Example: `[Name] IN ('Alice', 'Bob', 'Charlie')`

### `LIKE` and `LIKEI`

- Pattern matching with wildcards, evaluating to a boolean:
    * `_` matches a single character.
    * `%` matches zero or more characters.
- `LIKE` is case-sensitive, while `LIKEI` is case-insensitive.
- Example: `[Name] LIKE 'Ben%'`

### `NULL`

- Represents a missing value.
- Example: `[Age] IS NULL`

---

## Examples

### Filtering Rows

1. Filter rows where `Name` is `'Charlie'`:
    `[Name] == 'Charlie'`

2.	Filter rows where `Age` is greater than `30`:
    `[Age] > 30`

3.	Filter rows where `Name` starts with `'Ben'`:
    `[Name] LIKE 'Ben%'`
   

### Adding Fields

1. Add a field with a conditional value:
   `IIF([Age] > 30, 'Senior', 'Junior')`

2. Add a field with a formatted string:
   `'Name: ' + [Name] + ', Age: ' + [Age]`

3. Add a field with a cumulative sum by quarter:
   `CUMSUM([Sales], [Quarter])`