namespace XlBlocks.AddIn.Dna;

using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Registration;

internal static class RegistrationUtilities
{
    internal static IEnumerable<ExcelFunctionRegistration> GetExcelFunctions(bool includeTestFunctions)
    {
        return ExcelIntegration.GetExportedAssemblies()
            .SelectMany(x => x.GetTypes())
            .SelectMany(x => x.GetMethods(BindingFlags.Public | BindingFlags.Static))
            .Where(x => x.GetCustomAttribute<ExcelFunctionAttribute>() is not null)
            .Where(x => includeTestFunctions || x.GetCustomAttribute<IntegrationTestExcelFunctionAttribute>() is null)
            .Select(x => GetRegistration(x));
    }

    private static ExcelFunctionRegistration GetRegistration(MethodInfo methodInfo)
    {
        var functionRegistration = new ExcelFunctionRegistration(methodInfo);
        foreach (var parameterRegistration in functionRegistration.ParameterRegistrations)
        {
            var defaultAttribute = parameterRegistration.CustomAttributes.OfType<DefaultParameterValueAttribute>().FirstOrDefault();
            if (defaultAttribute is not null || parameterRegistration.CustomAttributes.OfType<OptionalAttribute>().Any())
            {
                parameterRegistration.ArgumentAttribute.Name = $"[{parameterRegistration.ArgumentAttribute.Name}]";

                var defaultStr = GetDefaultDescriptionString(defaultAttribute?.Value);
                parameterRegistration.ArgumentAttribute.Description = $"(Optional{defaultStr}) {parameterRegistration.ArgumentAttribute.Description}";
            }
        }
        return functionRegistration;
    }

    private static string? GetDefaultDescriptionString(object? defaultValue)
    {
        var defaultStr = defaultValue switch
        {
            string stringValue => $"'{stringValue}'",
            bool boolValue => $"{(boolValue ? "TRUE" : "FALSE")}",
            DateTime dateValue => $"{dateValue:yyyy-mm-dd}",
            Missing _ => string.Empty,
            _ => defaultValue?.ToString() ?? string.Empty
        };

        return string.IsNullOrEmpty(defaultStr) ? "" : $", default {defaultStr}";
    }
}
