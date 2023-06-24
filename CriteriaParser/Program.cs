using System.Collections.Immutable;
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
var paragraphs = documentBody!.ChildElements
    .Where(element => element is Paragraph paragraph && paragraph.ChildElements.Any(child => child is Run))
    .Cast<Paragraph>()
    .ToList();

var centeredParagraphs = paragraphs
    .Where(p => p.ChildElements.OfType<ParagraphProperties>()
        .Any(prop => prop.ChildElements.OfType<Justification>()
            .Any(j => j.Val!.Value == JustificationValues.Center)))
    .ToList();
    
var units = centeredParagraphs
    .Select(p => string.Join(string.Empty, p.ChildElements.OfType<Run>().Select(r => r.InnerText)))
    .Where(value => UnitRegex().Match(value).Success)
    .Select(value =>
    {
        var match = UnitRegex().Match(value);
        return new Unit
        {
            Text = value,
            Digit = int.Parse(match.Groups[1].Value)
        };
    })
    .OrderBy(unit => unit.Digit)
    .ToList();

var sections = centeredParagraphs
    .Where(p => p.ChildElements.OfType<ParagraphProperties>()
        .Any(pr => pr.ChildElements.OfType<ParagraphMarkRunProperties>()
            .Any(prop => prop.ChildElements.OfType<Italic>().Any())))
    .Select(p => string.Join(string.Empty, p.ChildElements.OfType<Run>().Select(r => r.InnerText)))
    .Where(value => SectionRegex().Match(value).Success)
    .Select(value =>
    {
        var match = SectionRegex().Match(value);
        return new Section
        {
            Text = value, 
            StartCriterionKey = int.Parse(match.Groups[1].Value),
            EndCriterionKey = int.Parse(match.Groups[2].Value)
        };
    })
    .OrderBy(section => section.StartCriterionKey)
    .ToList();

var criteriaDictionary = paragraphs
    .Where(p => p.ChildElements.OfType<Run>().Any(r => r.RunProperties!.ChildElements.Any(prop => prop is Bold)))
    .Select(p => string.Join(string.Empty, p.ChildElements.OfType<Run>().Select(r => r.InnerText)))
    .Where(value => CriterionRegex().Match(value).Success)
    .Select(value =>
    {
        var match = CriterionRegex().Match(value);
        return new Criterion
        {
            Key = int.Parse(match.Groups[1].Value),
            Text = value
        };
    })
    .ToImmutableSortedDictionary(criterion => criterion.Key, criterion => criterion);

foreach (var c in criteriaDictionary)
{
    var stringKey = c.Key.ToString();
    if (stringKey.Length > 3)
    {
        var parentKey = stringKey[..^1];
        if (criteriaDictionary.TryGetValue(int.Parse(parentKey), out var parent))
        {
            parent.Inner ??= new SortedDictionary<int, Criterion>();
            parent.Inner.Add(c.Key, c.Value);
        }
        else
        {
            Console.WriteLine($"No parent for '{parentKey}'");
        }
    }
}

foreach (var unit in units)
{
    unit.SectionList = sections.Where(section => section.StartCriterionKey / 100 == unit.Digit).ToList();
    
    // TODO: This approach loses unsectioned criteria in between. Fix later
    var unitFirstSectionedCriterion = unit.SectionList.FirstOrDefault()?.StartCriterionKey;   
    var unitLastSectionedCriterion = unit.SectionList.LastOrDefault()?.EndCriterionKey;
    foreach (var section in unit.SectionList)
    {
        section.CriterionList = criteriaDictionary
            .Where(pair => pair.Key >= section.StartCriterionKey
                           && pair.Key <= section.EndCriterionKey)
            .Select(pair => pair.Value)
            .ToList();
    }

    var unitHasSections = unitFirstSectionedCriterion.HasValue && unitLastSectionedCriterion.HasValue;

    unit.UnSectionedCriterionList = criteriaDictionary
        .Where(pair =>
            unitHasSections &&
            (pair.Key < unitFirstSectionedCriterion!.Value && pair.Key >= unit.Digit * 100 ||
             pair.Key > unitLastSectionedCriterion!.Value && pair.Key <= (unit.Digit + 1) * 100 - 1)
            || !unitHasSections && (pair.Key >= unit.Digit * 100 && pair.Key <= (unit.Digit + 1) * 100 - 1))
        .Select(pair => pair.Value)
        .ToList();
}

var options = new JsonSerializerOptions
{
    Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic),
    WriteIndented = true,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingDefault
};
var jsonResult = JsonSerializer.Serialize(units, options);
File.WriteAllText(outputFileName, jsonResult);

partial class Program
{
    [GeneratedRegex(@"Раздел \d\.(\d)\.(.*)")]
    private static partial Regex UnitRegex();
    
    [GeneratedRegex( @"\(d\s*(\d*)\s*\-\s*d\s*(\d*)")]
    private static partial Regex SectionRegex();
    
    [GeneratedRegex(@"d\s*(\d+)\ (.*)")]
    private static partial Regex CriterionRegex();
}