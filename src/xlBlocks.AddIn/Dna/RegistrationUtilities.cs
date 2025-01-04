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
                if (parameterRegistration.CustomAttributes.OfType<OptionalAttribute>().Any() || parameterRegistration.CustomAttributes.OfType<DefaultParameterValueAttribute>().Any())
                {
                    parameterRegistration.ArgumentAttribute.Name = $"[{parameterRegistration.ArgumentAttribute.Name}]";
                    parameterRegistration.ArgumentAttribute.Description = $"[Optional] {parameterRegistration.ArgumentAttribute.Description}";
                }
            }
        }
        return functionRegistrations;
    }
}
