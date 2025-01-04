namespace XlBlocks.AddIn.Dna;

using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelDna.Registration;

internal static class RegistrationUtilities
{
    internal static IEnumerable<ExcelFunctionRegistration> EnrichRegistrations(this IEnumerable<ExcelFunctionRegistration> functionRegistrations)
    {
        foreach (var functionRegistration in functionRegistrations)
        {
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
        }
        return functionRegistrations;
    }

    private static string? GetDefaultDescriptionString(object? defaultValue)
    {
        var defaultStr = defaultValue switch
        {
            string stringValue => $"'{stringValue}'",
            bool boolValue => $"{(boolValue ? "TRUE" : "FALSE")}",
            DateTime dateValue => $"{dateValue:yyyy-mm-dd}",
            _ => defaultValue?.ToString() ?? string.Empty
        };

        return string.IsNullOrEmpty(defaultStr) ? "" : $", default {defaultStr}";
    }
}
