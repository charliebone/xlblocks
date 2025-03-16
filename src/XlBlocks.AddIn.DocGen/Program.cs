using System.Reflection;
using ExcelDna.Integration;
using ExcelDna.Registration;
using Grynwald.MarkdownGenerator;
using Microsoft.Extensions.Configuration;
using XlBlocks.AddIn.Dna;

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
    .Select(x => new ExcelFunctionRegistration(x));

var functionsByType = excelFunctions
    .GroupBy(x => x.FunctionAttribute.Name.Split("_")[0]);

var excelDocSet = new DocumentSet<MdDocument>();
foreach (var functionGroup in functionsByType)
{
    // each group gets its own page
    var trimmedKey = functionGroup.Key.Replace("XB", "");
    var groupDocument = excelDocSet.CreateMdDocument($"{trimmedKey.ToLower()}.md");
    groupDocument.Root.Add(new MdHeading(1, trimmedKey));
}

// save to /docs/excel folder
excelDocSet.Save(outputPath);
Console.WriteLine("Doc gen is finished.");
