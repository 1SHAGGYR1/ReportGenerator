using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Shared;

namespace ReportGenerator;

internal class Program
{
    private const string DatesFormat = "dd.MM.yyyy";
    
    public static void Main(string[] args)
    {
        var criteriaFilePath = @"C:\Users\merkulov.e\Source\Playground\ReportGenerator\CriteriaParser\JsonView.json";
        var outputFileDirectory = string.Join('\\', Directory.GetCurrentDirectory().Split('\\')[..^3]);
        if (args.Any())
        {
            foreach (var argument in args)
            {
                switch (argument)
                {
                    case "criteriaFilePath":
                        criteriaFilePath = argument;
                        break;
                    case "outputFileDirectory":
                        outputFileDirectory = argument;
                        break;
                }
            }
        }

        var info = InputDocumentInfo();

        var outputFileName = outputFileDirectory +
                             $@"\функционирование_{info.ChildInfo.LastName}_{info.PeriodStartDate.ToString(DatesFormat)}-{info.PeriodEndDate.ToString(DatesFormat)}.docx";
        using var createdDocument = WordprocessingDocument.Create(outputFileName, WordprocessingDocumentType.Document);
        Console.WriteLine($"Создан документ {outputFileName}");

        var body = CreateDocumentBody(createdDocument);
        AddChildInfo(body, info.ChildInfo);
        var table = CreateReportTable(body);
        // var criteriaList = ParseCriteriaList(criteriaFilePath);
        Console.WriteLine("Приступаем к заполнению критериев");
        AddSectionProperties(body);
    }


    private static DocumentInfo InputDocumentInfo()
    {
        var result = new DocumentInfo();
        var lastNameIsValid = false;
        do
        {
            Console.WriteLine("Введите фамилию ребенка");
            var lastName = Console.ReadLine();
            if (string.IsNullOrEmpty(lastName))
            {
                Console.WriteLine("Фамилия не должна быть пустой");
            }
            else
            {
                lastNameIsValid = true;
                result.ChildInfo.LastName = lastName;
            }
        } while (!lastNameIsValid);

        var firstNameIsValid = false;
        do
        {
            Console.WriteLine("Введите имя ребенка");
            var firstName = Console.ReadLine();
            if (string.IsNullOrEmpty(firstName))
            {
                Console.WriteLine("Имя не должно быть пустым");
            }
            else
            {
                firstNameIsValid = true;
                result.ChildInfo.FirstName = firstName;
            }
        } while (!firstNameIsValid);

        var startDateIsValid = false;
        do
        {
            Console.WriteLine("Введите дату начала исследуемого периода в формате: 25.03.1999");
            var stringStartDate = Console.ReadLine();
            if (!DateOnly.TryParseExact(stringStartDate, DatesFormat, out var startDate))
            {
                Console.WriteLine("Неверный формат введенной даты.");
            }
            else
            {
                startDateIsValid = true;
                result.PeriodStartDate = startDate;
            }
        } while (!startDateIsValid);

        var endDateIsValid = false;
        do
        {
            Console.WriteLine("Введите дату конца исследуемого периода в формате: 15.03.1999");
            var stringEndDate = Console.ReadLine();
            if (!DateOnly.TryParseExact(stringEndDate, DatesFormat, out var endDate))
            {
                Console.WriteLine("Неверный формат введенной даты");
            }
            else
            {
                if (endDate < result.PeriodStartDate)
                {
                    Console.WriteLine("Дата конца периода должна быть больше даты начала");
                }
                else
                {
                    endDateIsValid = true;
                    result.PeriodEndDate = endDate;
                }
            }
        } while (!endDateIsValid);

        return result;
    }

    private static void AddChildInfo(Body body, ChildInfo info)
    {
        body.AppendChild(
            new Paragraph(
                new ParagraphProperties(
                    new Tabs(
                        new TabStop
                        {
                            Val = new EnumValue<TabStopValues>(TabStopValues.Left ),
                            Position = new Int32Value(5670)
                        })),
                new Run(
                    new RunProperties(new Bold()),
                    new Text($"{string.Concat(info.LastName[0].ToString().ToUpper(), info.LastName.AsSpan(1))} {string.Concat(info.FirstName[0].ToString().ToUpper(), info.FirstName.AsSpan(1))}"))
            )
        );
    }

    private static Body CreateDocumentBody(WordprocessingDocument createdDocument)
    {
        var body = new Body();
        createdDocument.AddMainDocumentPart().Document = new Document(body);

        return body;
    }

    private static Table CreateReportTable(Body body)
    {
        var table = new Table(
            new TableProperties(
                new TableWidth
                {
                    Width = new StringValue(DocumentMetrics.TableWidth),
                    Type = new EnumValue<TableWidthUnitValues>(TableWidthUnitValues.Dxa)
                },
                new TableBorders(
                    new TopBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = new UInt32Value((uint) 4),
                        Space = 0,
                        Color = new StringValue("auto")
                    },
                    new LeftBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = new UInt32Value((uint) 4),
                        Space = 0,
                        Color = new StringValue("auto")
                    },
                    new BottomBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = new UInt32Value((uint) 4),
                        Space = 0,
                        Color = new StringValue("auto")
                    },
                    new RightBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = new UInt32Value((uint) 4),
                        Space = 0,
                        Color = new StringValue("auto")
                    },
                    new InsideHorizontalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = new UInt32Value((uint) 4),
                        Space = 0,
                        Color = new StringValue("auto")
                    },
                    new InsideVerticalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = new UInt32Value((uint) 4),
                        Space = 0,
                        Color = new StringValue("auto")
                    })),
            new TableGrid(
                new GridColumn { Width = new StringValue(DocumentMetrics.TableCellWidth) },
                new GridColumn { Width = new StringValue(DocumentMetrics.TableCellWidth) }),
            new TableRow(
                new TableCell(
                    new TableCellProperties(
                        new TableCellWidth { Width = DocumentMetrics.TableWidth },
                        new GridSpan{Val = new Int32Value(2)}),
                    new Paragraph(
                        new ParagraphProperties(
                            new Justification {Val = new EnumValue<JustificationValues>(JustificationValues.Center)}),
                        new Run(
                            new RunProperties(
                                new Bold(),
                                new FontSize{Val = new StringValue("32")},
                                new FontSizeComplexScript{Val = new StringValue("32")}),
                            new Text("Таблица критериев"))))));
        body.AppendChild(table);

        return table;
    }

    private static List<Criteria> ParseCriteriaList(string path)
    {
        try
        {
            var jsonString = File.ReadAllText(path);
            return JsonSerializer.Deserialize<List<Criteria>>(jsonString)!;
        }
        catch (Exception e)
        {
            Console.WriteLine("Ошибка десериализации критериев");
            Console.WriteLine(e);
            throw;
        }
    }
    
    private static void AddSectionProperties(Body body)
    {
        body.AppendChild(new SectionProperties(
            new PageSize
            {
                Width = new UInt32Value(DocumentMetrics.PageSize),
                Orient = new EnumValue<PageOrientationValues>(PageOrientationValues.Landscape)
            },
            new PageMargin
            {
                Top = new Int32Value(DocumentMetrics.PageMargins.Top),
                Right = new UInt32Value(DocumentMetrics.PageMargins.Right),
                Bottom = new Int32Value(DocumentMetrics.PageMargins.Bottom),
                Left = new UInt32Value(DocumentMetrics.PageMargins.Left),
                Header = new UInt32Value(DocumentMetrics.PageMargins.Header),
                Footer = new UInt32Value(DocumentMetrics.PageMargins.Footer),
            }));
    }
}