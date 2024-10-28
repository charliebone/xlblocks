namespace XlBlocks.AddIn.Parser;

using System.Collections.Generic;
using NLog;
using sly.lexer;
using sly.parser;
using sly.parser.generator;
using sly.parser.parser;
using XlBlocks.AddIn.Parser.Expressions;

internal class DataFrameExpressionParser
{
    private static readonly Logger _logger = LogManager.GetLogger(typeof(DataFrameExpressionParser).FullName);

    public static readonly Parser<DataFrameExpressionToken, IColumnExpression> Parser = GetParser();

#pragma warning disable CA1822

    private static Parser<DataFrameExpressionToken, IColumnExpression> GetParser()
    {
        // parser should be thread-safe so long as context visitor is safe: https://github.com/b3b00/csly/issues/182
        var parserInstance = new DataFrameExpressionParser();
        var builder = new ParserBuilder<DataFrameExpressionToken, IColumnExpression>();
        var parser = builder.BuildParser(parserInstance, ParserType.EBNF_LL_RECURSIVE_DESCENT, $"{nameof(DataFrameExpressionParser)}_expressions");
        if (parser.IsError)
        {
            var message = $"errors when building {nameof(DataFrameExpressionParser)}: '{string.Join("', '", parser.Errors.Select(x => x.Message))}'";
            _logger.Error(message);
            throw new Exception(message);
        }

        return parser.Result;
    }

    [Prefix((int)DataFrameExpressionToken.NOT, Associativity.Right, 90)]
    [Prefix((int)DataFrameExpressionToken.EXCLAMATION_POINT, Associativity.Right, 90)]
    [Prefix((int)DataFrameExpressionToken.ARITH_MINUS, Associativity.Right, 29)]
    public IColumnExpression Operator_NegatePrefix(Token<DataFrameExpressionToken> negationToken, IColumnExpression expression)
    {
        return new UnaryColumnExpression(negationToken, expression);
    }

    [Infix((int)DataFrameExpressionToken.IN, Associativity.Left, 70)]
    [Infix((int)DataFrameExpressionToken.LIKE, Associativity.Left, 70)]
    public IColumnExpression Operator_InLike(IColumnExpression left, Token<DataFrameExpressionToken> likeOrInToken, IColumnExpression right)
    {
        return likeOrInToken.TokenID == DataFrameExpressionToken.LIKE ?
            new LikeClauseExpression(left, right) :
            new InClauseExpression(left, right);
    }

    [Infix((int)DataFrameExpressionToken.CARET, Associativity.Left, 50)]

    [Infix((int)DataFrameExpressionToken.ARITH_TIMES, Associativity.Left, 40)]
    [Infix((int)DataFrameExpressionToken.ARITH_DIVIDE, Associativity.Left, 40)]

    [Infix((int)DataFrameExpressionToken.ARITH_MODULO, Associativity.Left, 35)]

    [Infix((int)DataFrameExpressionToken.ARITH_PLUS, Associativity.Left, 30)]
    [Infix((int)DataFrameExpressionToken.ARITH_MINUS, Associativity.Left, 30)]

    [Infix((int)DataFrameExpressionToken.COMP_EQUALS, Associativity.Left, 20)]
    [Infix((int)DataFrameExpressionToken.COMP_NOTEQUALS, Associativity.Left, 20)]
    [Infix((int)DataFrameExpressionToken.COMP_LT, Associativity.Left, 20)]
    [Infix((int)DataFrameExpressionToken.COMP_GT, Associativity.Left, 20)]
    [Infix((int)DataFrameExpressionToken.COMP_LTE, Associativity.Left, 20)]
    [Infix((int)DataFrameExpressionToken.COMP_GTE, Associativity.Left, 20)]

    [Infix((int)DataFrameExpressionToken.IS, Associativity.Left, 15)]

    [Infix((int)DataFrameExpressionToken.AND, Associativity.Left, 11)]

    [Infix((int)DataFrameExpressionToken.OR, Associativity.Left, 10)]
    [Infix((int)DataFrameExpressionToken.XOR, Associativity.Left, 10)]
    public IColumnExpression Operator_Binary(IColumnExpression left, Token<DataFrameExpressionToken> arithmeticOpToken, IColumnExpression right)
    {
        return new BinaryColumnExpression(left, arithmeticOpToken, right);
    }

    [Operand]
    [Production("primary : BRACKET_LEFT [d] IDENTIFIER BRACKET_RIGHT [d]")]
    public IColumnExpression Primary_BracketedIdentifier(Token<DataFrameExpressionToken> identifier)
    {
        return new ColumnExpression(identifier.StringWithoutQuotes);
    }

    [Operand]
    [Production($"primary : PARENS_LEFT [d] {nameof(DataFrameExpressionParser)}_expressions PARENS_RIGHT [d]")]
    public IColumnExpression Primary_OperandGroup(IColumnExpression expression)
    {
        return expression;
    }

    [Operand]
    [Production("literal : INT")]
    public IColumnExpression Literal_INT(Token<DataFrameExpressionToken> literal)
    {
        return new ConstantColumnExpression<int>(literal.IntValue);
    }

    [Operand]
    [Production("literal : NUMBER")]
    public IColumnExpression Literal_NUMBER(Token<DataFrameExpressionToken> literal)
    {
        return new ConstantColumnExpression<double>(literal.DoubleValue);
    }

    [Operand]
    [Production("literal : STRING")]
    public IColumnExpression Literal_STRING(Token<DataFrameExpressionToken> literal)
    {
        return new ConstantColumnExpression<string>(literal.StringWithoutQuotes);
    }

    [Operand]
    [Production("literal : [ TRUE | FALSE ]")]
    public IColumnExpression Literal_BOOLEAN(Token<DataFrameExpressionToken> literal)
    {
        return new ConstantColumnExpression<bool>(literal.TokenID == DataFrameExpressionToken.TRUE);
    }

    [Production("literal : NULL [d]")]
    public IColumnExpression Literal_NULL()
    {
        return new NullExpression();
    }

    [Operand]
    [Production("literal_list : PARENS_LEFT [d] literal (COMMA [d] literal)* PARENS_RIGHT [d]")]
    public IColumnExpression Literal_List(IColumnExpression firstLiteral, List<Group<DataFrameExpressionToken, IColumnExpression>> restListerals)
    {
        return new ExpressionListExpression(firstLiteral, restListerals);
    }

    [Operand]
    [Production("function_call : IDENTIFIER PARENS_LEFT [d] operand_list PARENS_RIGHT [d]")]
    public IColumnExpression Function_Call(Token<DataFrameExpressionToken> identifier, IColumnExpression expressionList)
    {
        return new FunctionCallExpression(identifier, expressionList);
    }

    [Production($"operand_list : DataFrameExpressionParser_expressions (COMMA [d] {nameof(DataFrameExpressionParser)}_expressions)* ")]
    public IColumnExpression Operand_List(IColumnExpression firstExpression, List<Group<DataFrameExpressionToken, IColumnExpression>> restExpressions)
    {
        return new ExpressionListExpression(firstExpression, restExpressions);
    }
}
