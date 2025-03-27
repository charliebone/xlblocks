namespace XlBlocks.AddIn.Parser;

using sly.i18n;
using sly.lexer;

[Lexer]
public enum DataFrameExpressionToken
{
    [Lexeme(GenericToken.Extension)]
    [LexemeLabel("en", "column name")]
    COLUMN_IDENTIFIER,

    [LexemeLabel("en", "function name")]
    [Lexeme(GenericToken.Identifier, IdentifierType.Custom, "_A-Za-z\\.", "-_0-9A-Za-z\\.")]
    FUNCTION_IDENTIFIER,

    [LexemeLabel("en", "integer")]
    [Int]
    INT,

    [LexemeLabel("en", "number")]
    [Double]
    NUMBER,

    [LexemeLabel("en", "string")]
    [String("\"", "\\")]
    [String("\'", "\\")]
    STRING,

    [LexemeLabel("en", "NOT")]
    [Keyword("NOT")]
    [Keyword("not")]
    NOT,

    [LexemeLabel("en", "AND")]
    [Keyword("AND")]
    [Keyword("and")]
    AND,

    [LexemeLabel("en", "OR")]
    [Keyword("OR")]
    [Keyword("or")]
    OR,

    [LexemeLabel("en", "XOR")]
    [Keyword("XOR")]
    [Keyword("xor")]
    XOR,

    [LexemeLabel("en", "IS")]
    [Keyword("IS")]
    [Keyword("is")]
    IS,

    [LexemeLabel("en", "NULL")]
    [Keyword("NULL")]
    [Keyword("null")]
    NULL,

    [LexemeLabel("en", "IN")]
    [Keyword("IN")]
    [Keyword("in")]
    IN,

    [LexemeLabel("en", "LIKE")]
    [Keyword("LIKE")]
    [Keyword("like")]
    LIKE,

    [LexemeLabel("en", "LIKEI")]
    [Keyword("LIKEI")]
    [Keyword("likei")]
    LIKEI,

    [LexemeLabel("en", "TRUE")]
    [Keyword("TRUE")]
    [Keyword("true")]
    TRUE,

    [LexemeLabel("en", "FALSE")]
    [Keyword("FALSE")]
    [Keyword("false")]
    FALSE,

    [LexemeLabel("en", "exclamation point")]
    [Sugar("!")]
    EXCLAMATION_POINT,

    [LexemeLabel("en", "caret")]
    [Sugar("^")]
    CARET,

    [LexemeLabel("en", "plus")]
    [Sugar("+")]
    ARITH_PLUS,

    [LexemeLabel("en", "minus")]
    [Sugar("-")]
    ARITH_MINUS,

    [LexemeLabel("en", "multiplication")]
    [Sugar("*")]
    ARITH_TIMES,

    [LexemeLabel("en", "division")]
    [Sugar("/")]
    ARITH_DIVIDE,

    [LexemeLabel("en", "modulo")]
    [Sugar("%")]
    ARITH_MODULO,

    [LexemeLabel("en", "equal to")]
    [Sugar("==")]
    COMP_EQUALS,

    [LexemeLabel("en", "not equal to")]
    [Sugar("<>")]
    [Sugar("!=")]
    COMP_NOTEQUALS,

    [LexemeLabel("en", "less than")]
    [Sugar("<")]
    COMP_LT,

    [LexemeLabel("en", "greater than")]
    [Sugar(">")]
    COMP_GT,

    [LexemeLabel("en", "less than or equal to")]
    [Sugar("<=")]
    COMP_LTE,

    [LexemeLabel("en", "greater than or equal to")]
    [Sugar(">=")]
    COMP_GTE,

    [LexemeLabel("en", "comma")]
    [Sugar(",")]
    COMMA,

    [LexemeLabel("en", "left parenthesis")]
    [Sugar("(")]
    PARENS_LEFT,

    [LexemeLabel("en", "right parenthesis")]
    [Sugar(")")]
    PARENS_RIGHT,
}
