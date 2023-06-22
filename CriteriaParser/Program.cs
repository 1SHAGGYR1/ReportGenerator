using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Text.Unicode;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Shared;

const string inputFileName = @"C:\Users\merkulov.e\Source\Playground\ReportGenerator\CriteriaParser\активность и участие.docx";
const string outputFileName = @"C:\Users\merkulov.e\Source\Playground\ReportGenerator\CriteriaParser\JsonView.json"; 

using var myDocument = WordprocessingDocument.Open(inputFileName, false);
var documentBody = myDocument.MainDocumentPart!.Document.Body;
var criteriaDictionary = documentBody!.ChildElements
    .Where(element => element is Paragraph paragraph && paragraph.ChildElements.Any(child => child is Run))
    .Cast<Paragraph>()
    .Where(p => p.ChildElements.OfType<Run>().Any(r => r.RunProperties!.ChildElements.Any(prop => prop is Bold)))
    .Select(p => string.Join(string.Empty, p.ChildElements.OfType<Run>().Select(r => r.InnerText)))
    .Where(value => MyRegex().Match(value).Success)
    .Select(value =>
    {
        var match = MyRegex().Match(value);
        return new Criteria {Key = match.Groups[1].Value.Replace(" ", string.Empty), Text = match.Groups[2].Value.Trim()};
    })
    .ToDictionary(criteria => criteria.Key, criteria => criteria);

foreach (var c in criteriaDictionary)
{
    if (c.Key.Length > 4)
    {
        var parentKey = c.Key[..^1];
        if (criteriaDictionary.TryGetValue(parentKey, out var parent))
        {
            parent.Inner ??= new Dictionary<string, Criteria>();
            parent.Inner.Add(c.Key, c.Value);
        }
        else
        {
            Console.WriteLine($"No parent for '{parentKey}'");
        }
    }
}

var result = criteriaDictionary.Values.Where(item => item.Key.Length == 4).ToList();
var options = new JsonSerializerOptions
{
    Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic),
    WriteIndented = true,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingDefault
};
var jsonResult = JsonSerializer.Serialize(result, options);
File.WriteAllText(outputFileName, jsonResult);

partial class Program
{
    [GeneratedRegex(@"(d\ \d+)\ (.*)")]
    private static partial Regex MyRegex();
}