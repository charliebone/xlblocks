using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Registration;
using Grynwald.MarkdownGenerator;
using Microsoft.Extensions.Configuration;
using XlBlocks.AddIn.Dna;
using XlBlocks.AddIn.DocGen;

// generate documentation based on excel functions
Console.WriteLine("Starting doc gen...");

var config = new ConfigurationBuilder()
    .AddCommandLine(args)
    .Build();

var outputPath = config["OutputPath"] ?? "../../../../docs/docs/excel";
if (!Directory.Exists(outputPath))
    Directory.CreateDirectory(outputPath);

var addInAssembly = Assembly.GetAssembly(typeof(XlBlocksAddIn)) ?? throw new Exception("cannot find AddIn assembly");
var excelFunctions = addInAssembly.GetTypes()
    .SelectMany(x => x.GetMethods(BindingFlags.Public | BindingFlags.Static))
    .Where(x => x.GetCustomAttribute<ExcelFunctionAttribute>() is not null)
    .Where(x => x.GetCustomAttribute<IntegrationTestExcelFunctionAttribute>() is null)
    .Select(x => new{ Function = x, Registration = new ExcelFunctionRegistration(x) });

var functionsByType = excelFunctions
    .GroupBy(x => x.Registration.FunctionAttribute.Name.Split("_")[0]);

var excelDocSet = new DocumentSet<MdDocument>();
foreach (var functionGroup in functionsByType)
{
    // each group gets its own page
    //var trimmedKey = functionGroup.Key.Replace("XB", "");
    var groupDocument = excelDocSet.CreateMdDocument($"{functionGroup.Key.Replace("XB", "").ToLower()}.md");

    //var title = trimmedKey == "Dict" ? "Dictionary" : trimmedKey;
    //groupDocument.Root.Add(new MdHeading(1, title));

    foreach (var function in functionGroup)
    {
        groupDocument.Root.Add(new MdHeading(3, function.Registration.FunctionAttribute.Name));
        groupDocument.Root.Add(new MdParagraph(function.Registration.FunctionAttribute.Description));

        groupDocument.Root.Add(new MdCodeBlock(function.Registration.GetSignature(), "excel"));

        var paramTable = new List<MdTableRow>();
        var lambdaParameters = function.Registration.FunctionLambda.Parameters.Where(x => x.Name is not null).ToDictionary(x => x.Name?.ToLower() ?? "", x => x.Type);
        foreach (var parameter in function.Registration.ParameterRegistrations)
        {
            var defaultAttribute = parameter.CustomAttributes.OfType<DefaultParameterValueAttribute>().FirstOrDefault();
            var isOptional = defaultAttribute is not null || parameter.CustomAttributes.OfType<OptionalAttribute>().Any();

            var parameterType = lambdaParameters.ContainsKey(parameter.ArgumentAttribute.Name.ToLower()) ? 
                lambdaParameters[parameter.ArgumentAttribute.Name.ToLower()].Name : "Any";

            var rowSpans = new List<MdSpan>
            {
                new MdTextSpan(parameter.ArgumentAttribute.Name),
                new MdCodeSpan(parameterType)
            };
            if (isOptional)
            {
                rowSpans.Add(new MdCompositeSpan(
                    new MdEmphasisSpan("Optional  "),
                    new MdTextSpan($"{parameter.ArgumentAttribute.Description}")));
            }
            else
            {
                rowSpans.Add(new MdTextSpan(parameter.ArgumentAttribute.Description));
            }

            paramTable.Add(new MdTableRow(rowSpans));
        }

        if (paramTable.Count > 0)
            groupDocument.Root.Add(new MdTable(new MdTableRow(new[] { "Parameter", "Type", "Description" }), paramTable));

        groupDocument.Root.Add(new MdHeading(6, "Returns"));
        groupDocument.Root.Add(new MdParagraph(function.Function.ReturnType.GetReturnTypeName()));

    }
}

// save to /docs/excel folder
excelDocSet.Save(outputPath, false, (MdDocument document, string path) =>
{
    var trimmedKey = Path.GetFileNameWithoutExtension(path);

    // write generated content
    document.Save(path, MdSerializationOptions.Presets.MkDocs);

    // append base content
    var baseFile = Path.Combine(outputPath, "../excel-base", $"{trimmedKey.ToLower()}.md");
    if (File.Exists(baseFile))
    {
        var baseContents = File.ReadAllText(baseFile);
        var generatedContents = File.ReadAllText(path);
        var combinedContent = baseContents + Environment.NewLine + generatedContents;
        File.WriteAllText(path, combinedContent);
    }
});

Console.WriteLine("Doc gen is finished.");
