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

    [Prefix((int)DataFrameExpressionToken.NOT, Associativity.Right, 92)]
    [Prefix((int)DataFrameExpressionToken.EXCLAMATION_POINT, Associativity.Right, 91)]
    public IColumnExpression Expression_Not_Expression(Token<DataFrameExpressionToken> negationToken, IColumnExpression expression)
    {
        return new UnaryColumnExpression(negationToken, expression);
    }

    [Infix((int)DataFrameExpressionToken.IN, Associativity.Right, 71)]
    [Infix((int)DataFrameExpressionToken.LIKE, Associativity.Right, 70)]
    public IColumnExpression Expression_IN_Clause(IColumnExpression left, Token<DataFrameExpressionToken> likeOrInToken, IColumnExpression right)
    {
        return likeOrInToken.TokenID == DataFrameExpressionToken.LIKE ?
            new LikeClauseExpression(left, right) :
            new InClauseExpression(left, right);
    }

    [Infix((int)DataFrameExpressionToken.ARITH_TIMES, Associativity.Left, 40)]
    [Infix((int)DataFrameExpressionToken.ARITH_DIVIDE, Associativity.Left, 40)]
    [Infix((int)DataFrameExpressionToken.ARITH_MODULO, Associativity.Left, 38)]
    [Infix((int)DataFrameExpressionToken.ARITH_PLUS, Associativity.Left, 30)]
    [Infix((int)DataFrameExpressionToken.ARITH_MINUS, Associativity.Left, 30)]
    public IColumnExpression Expression_AS_Expression(IColumnExpression left, Token<DataFrameExpressionToken> arithmeticOpToken, IColumnExpression right)
    {
        return new BinaryColumnExpression(left, arithmeticOpToken, right);
    }

    [Prefix((int)DataFrameExpressionToken.ARITH_MINUS, Associativity.Right, 29)]
    public IColumnExpression Expression_Negative_Expression(Token<DataFrameExpressionToken> negativeToken, IColumnExpression right)
    {
        return new UnaryColumnExpression(negativeToken, right);
    }

    [Infix((int)DataFrameExpressionToken.COMP_EQUALS, Associativity.Left, 25)]
    [Infix((int)DataFrameExpressionToken.COMP_NOTEQUALS, Associativity.Left, 24)]
    [Infix((int)DataFrameExpressionToken.COMP_LT, Associativity.Left, 23)]
    [Infix((int)DataFrameExpressionToken.COMP_GT, Associativity.Left, 22)]
    [Infix((int)DataFrameExpressionToken.COMP_LTE, Associativity.Left, 21)]
    [Infix((int)DataFrameExpressionToken.COMP_GTE, Associativity.Left, 20)]
    public IColumnExpression Expression_Expression_AndOr_Expression(IColumnExpression left, Token<DataFrameExpressionToken> comparisonOpToken, IColumnExpression right)
    {
        return new BinaryColumnExpression(left, comparisonOpToken, right);
    }

    [Infix((int)DataFrameExpressionToken.IS, Associativity.Left, 15)]
    public IColumnExpression Expression_Is(IColumnExpression left, Token<DataFrameExpressionToken> logicalOpToken, IColumnExpression right)
    {
        return new BinaryColumnExpression(left, logicalOpToken, right);
    }

    [Infix((int)DataFrameExpressionToken.AND, Associativity.Left, 12)]
    [Infix((int)DataFrameExpressionToken.OR, Associativity.Left, 11)]
    [Infix((int)DataFrameExpressionToken.XOR, Associativity.Left, 10)]
    public IColumnExpression Expression_AND_Expression(IColumnExpression left, Token<DataFrameExpressionToken> logicalOpToken, IColumnExpression right)
    {
        return new BinaryColumnExpression(left, logicalOpToken, right);
    }

    [Production("expression : DataFrameExpressionParser_expressions")]
    public IColumnExpression Expression_Operands(IColumnExpression operandExpression)
    {
        return operandExpression;
    }

    [Operand]
    [Production("primary : function_call")]
    public IColumnExpression Primary_FunctionCall(IColumnExpression functionCallExpression)
    {
        return functionCallExpression;
    }

    [Operand]
    [Production("primary : BRACKET_LEFT [d] IDENTIFIER BRACKET_RIGHT [d]")]
    public IColumnExpression ColumnName_Bracketed(Token<DataFrameExpressionToken> identifier)
    {
        return new ColumnExpression(identifier.StringWithoutQuotes);
    }

    [Operand]
    [Production("primary : PARENS_LEFT [d] expression PARENS_RIGHT [d]")]
    public IColumnExpression Primary_ExpressionGroup(IColumnExpression expression)
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

    [Production("function_call : IDENTIFIER PARENS_LEFT [d] expression_list? PARENS_RIGHT [d]")]
    public IColumnExpression FunctionCall(Token<DataFrameExpressionToken> identifier, ValueOption<IColumnExpression> expressionList)
    {
        return new FunctionCallExpression(identifier, expressionList);
    }

    /*[Production("expression : DataFrameExpressionParser_expressions IN [d] PARENS_LEFT [d] literal_list PARENS_RIGHT [d]")]
    public IColumnExpression InClause(IColumnExpression expression, IColumnExpression inList)
    {
        return new InClauseExpression(expression, inList);
    }*/

    [Production("expression_list : expression (COMMA [d] expression)* ")]
    public IColumnExpression ExpressionList(IColumnExpression firstExpression, List<Group<DataFrameExpressionToken, IColumnExpression>> restExpressions)
    {
        return new ExpressionListExpression(firstExpression, restExpressions);
    }

    [Operand]
    [Production("literal_list : literal (COMMA [d] literal)* ")]
    public IColumnExpression LiteralList(IColumnExpression firstLiteral, List<Group<DataFrameExpressionToken, IColumnExpression>> restListerals)
    {
        return new ExpressionListExpression(firstLiteral, restListerals);
    }
}
