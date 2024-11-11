namespace XlBlocks.AddIn.Parser;

using sly.i18n;
using sly.lexer;

[Lexer]
public enum DataFrameExpressionToken
{
    // note: hard requirement for columns starting only with letter, dot or underscore -- no leading numbers
    [LexemeLabel("en", "identifier")]
    [Lexeme(GenericToken.Identifier, IdentifierType.Custom, "_A-Za-z\\.", "-_0-9A-Za-z\\.")]
    IDENTIFIER,

    [LexemeLabel("en", "integer literal")]
    [Int]
    INT,

    [LexemeLabel("en", "numeric literal")]
    [Double]
    NUMBER,

    [LexemeLabel("en", "string literal")]
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

    [LexemeLabel("en", "is")]
    [Keyword("IS")]
    [Keyword("is")]
    IS,

    [LexemeLabel("en", "NULL")]
    [Keyword("NULL")]
    [Keyword("null")]
    NULL,

    [LexemeLabel("en", "in")]
    [Keyword("IN")]
    [Keyword("in")]
    IN,

    [LexemeLabel("en", "like")]
    [Keyword("LIKE")]
    [Keyword("like")]
    LIKE,

    [LexemeLabel("en", "case-insensitive like")]
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

    [LexemeLabel("en", "plus sign")]
    [Sugar("+")]
    ARITH_PLUS,

    [LexemeLabel("en", "minus sign")]
    [Sugar("-")]
    ARITH_MINUS,

    [LexemeLabel("en", "multiplication symbol")]
    [Sugar("*")]
    ARITH_TIMES,

    [LexemeLabel("en", "division symbol")]
    [Sugar("/")]
    ARITH_DIVIDE,

    [LexemeLabel("en", "modulo sign")]
    [Sugar("%")]
    ARITH_MODULO,

    [LexemeLabel("en", "is equal to")]
    [Sugar("==")]
    COMP_EQUALS,

    [LexemeLabel("en", "is not equal to")]
    [Sugar("<>")]
    [Sugar("!=")]
    COMP_NOTEQUALS,

    [LexemeLabel("en", "is less than")]
    [Sugar("<")]
    COMP_LT,

    [LexemeLabel("en", "is greater than")]
    [Sugar(">")]
    COMP_GT,

    [LexemeLabel("en", "is less than or equal to")]
    [Sugar("<=")]
    COMP_LTE,

    [LexemeLabel("en", "is greater than or equal to")]
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

    [LexemeLabel("en", "left bracket")]
    [Sugar("[")]
    BRACKET_LEFT,

    [LexemeLabel("en", "right bracket")]
    [Sugar("]")]
    BRACKET_RIGHT
}
