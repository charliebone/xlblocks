using System.Runtime.InteropServices;
using ExcelDna.Registration;

namespace XlBlocks.AddIn.DocGen
{
    public static class Helpers
    {
        public static string GetSignature(this ExcelFunctionRegistration function)
        {
            var argStr = string.Empty;
            foreach (var parameterRegistration in function.ParameterRegistrations)
            {
                var defaultAttribute = parameterRegistration.CustomAttributes.OfType<DefaultParameterValueAttribute>().FirstOrDefault();
                var isOptional = defaultAttribute is not null || parameterRegistration.CustomAttributes.OfType<OptionalAttribute>().Any();

                if (argStr.Length > 0)
                    argStr += ", ";

                argStr += $"{(isOptional ? "[" : "")}{parameterRegistration.ArgumentAttribute.Name}{(isOptional ? "]" : "")}";
            }
            return $"={function.FunctionAttribute.Name}({argStr})";
        }

        public static string GetReturnTypeName(this Type returnType)
        {
            return returnType.Name switch
            {
                "String" => "A string",
                "Double" => "A double",
                "Boolean" => "A boolean",
                "Object" => "An object",
                "Int32" => "An integer",
                "Int64" => "An integer",
                "DateTime" => "A date",
                "Object[]" => "A range",
                "Object[,]" => "A range",
                "XlBlockRange" => "A range",
                "XlBlockDictionary" => "A dictionary",
                "XlBlockList" => "A list",
                "XlBlockTable" => "A table",
                _ => returnType.Name
            };
        }
    }
}
