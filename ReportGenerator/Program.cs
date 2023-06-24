using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ReportGenerator.Models;
using Shared;

namespace ReportGenerator;

internal class Program
{
    private const string DatesFormat = "dd.MM.yyyy";
    
    //TODO: add ability to continue filling existing document
    public static void Main(string[] args)
    {
        var unitsFilePath = @"C:\Users\merkulov.e\Source\Playground\ReportGenerator\CriteriaParser\JsonView.json";
        var outputFileDirectory = string.Join('\\', Directory.GetCurrentDirectory().Split('\\')[..^3]);
        if (args.Any())
        {
            foreach (var argument in args)
            {
                switch (argument)
                {
                    case "unitsFilePath":
                        unitsFilePath = argument;
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
        AddSectionProperties(body);
        AddChildInfo(body, info.ChildInfo);
        var table = CreateReportTable(body);
        var unitsList = ParseUnitsList(unitsFilePath);
        Console.WriteLine("Приступаем к заполнению.");
        FillTable(unitsList, table);
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
                                new FontSize{Val = new StringValue(DocumentMetrics.Fonts.HeaderFontSize)},
                                new FontSizeComplexScript{Val = new StringValue(DocumentMetrics.Fonts.HeaderFontSize)}),
                            new Text("Таблица критериев"))))));
        body.AppendChild(table);

        return table;
    }

    private static List<Unit> ParseUnitsList(string path)
    {
        try
        {
            var jsonString = File.ReadAllText(path);
            return JsonSerializer.Deserialize<List<Unit>>(jsonString)!;
        }
        catch (Exception e)
        {
            Console.WriteLine("Ошибка десериализации разделов:");
            Console.WriteLine(e);
            throw;
        }
    }
    
    private static void FillTable(List<Unit> unitList, Table table)
    {
        try
        {
            foreach (var unit in unitList)
            {
                if (!InputUnit(unit)) continue;
                AddUnitRow(table, unit);

                foreach (var section in unit.SectionList)
                {
                    if (!InputSection(section)) continue;
                    AddSectionRow(table, section);

                    foreach (var criterion in section.CriterionList)
                    {
                        var criterionAnswer = InputCriterionAnswer(criterion);
                        Console.WriteLine();
                        AddCriterionRow(table, criterion, criterionAnswer);
                    }
                }

                foreach (var unSectionedCriterion in unit.UnSectionedCriterionList)
                {
                    var criterionAnswer = InputCriterionAnswer(unSectionedCriterion);
                    Console.WriteLine();
                    AddCriterionRow(table, unSectionedCriterion, criterionAnswer);
                }
            }
        }
        catch (FinishFillingException)
        {
        }
        finally
        {
            Console.WriteLine("Заполнение докунмента закончено.");
        }
    }
    
    private static bool InputUnit(Unit unit)
    {
        Console.WriteLine(OutputStrings.StartFillingPartMessage, "раздела", unit.Text);
        bool? inputUnit;
        do
        {
            Console.WriteLine(OutputStrings.SkipOption, "раздела");
            Console.WriteLine(OutputStrings.FinishOption);
            var input = Console.ReadKey();
            inputUnit = input.Key switch
            {
                ConsoleKey.D0 => false,
                ConsoleKey.Enter => true,
                ConsoleKey.Escape => throw new FinishFillingException(),
                _ => null
            };
            Console.WriteLine();
        } while (!inputUnit.HasValue);

        return inputUnit.Value;
    }

    private static bool InputSection(Section section)
    {
        Console.WriteLine(OutputStrings.StartFillingPartMessage, "секции", section.Text);
        bool? inputUnit;
        do
        {
            Console.WriteLine(OutputStrings.SkipOption, "секции");
            Console.WriteLine(OutputStrings.FinishOption);
            var input = Console.ReadKey();
            inputUnit = input.Key switch
            {
                ConsoleKey.D0 => false,
                ConsoleKey.Enter => true,
                ConsoleKey.Escape => throw new FinishFillingException(),
                _ => null
            };
            Console.WriteLine();
        } while (!inputUnit.HasValue);

        return inputUnit.Value;;
    }

    private static void AddUnitRow(Table table, Unit unit)
    {
        table.AppendChild(new TableRow(
            new TableCell(
                new TableCellProperties(
                    new TableCellWidth {Width = DocumentMetrics.TableWidth},
                    new GridSpan {Val = new Int32Value(2)}),
                new Paragraph(
                    new ParagraphProperties(
                        new Justification {Val = new EnumValue<JustificationValues>(JustificationValues.Center)}),
                    new Run(
                        new RunProperties(
                            new Bold(),
                            new FontSize {Val = new StringValue(DocumentMetrics.Fonts.UnitFontSize)},
                            new FontSizeComplexScript {Val = new StringValue(DocumentMetrics.Fonts.UnitFontSize)}),
                        new Text(unit.Text))))));
    }

    private static void AddSectionRow(Table table, Section section)
    {
        table.AppendChild(new TableRow(
            new TableCell(
                new TableCellProperties(
                    new TableCellWidth {Width = DocumentMetrics.TableWidth},
                    new GridSpan {Val = new Int32Value(2)}),
                new Paragraph(
                    new ParagraphProperties(
                        new Justification {Val = new EnumValue<JustificationValues>(JustificationValues.Center)}),
                    new Run(
                        new RunProperties(
                            new Italic(),
                            new FontSize {Val = new StringValue(DocumentMetrics.Fonts.SectionFontSize)},
                            new FontSizeComplexScript {Val = new StringValue(DocumentMetrics.Fonts.SectionFontSize)}),
                        new Text(section.Text))))));
    }

    private static string InputCriterionAnswer(Criterion criterion)
    {
        const string criterionAnswerTemplate =
            $"""
                Введите уровень затруднений:
                    1 - {OutputStrings.NoDifficulties}
                    2 - {OutputStrings.SmallDifficulties}
                    3 - {OutputStrings.MiddleDifficulties}
                    4 - {OutputStrings.HardDifficulties}
                    5 - {OutputStrings.TotalDifficulties}
            """;
        string criterionAnswer;
        do
        {
            Console.WriteLine(OutputStrings.StartFillingPartMessage, "критертия", criterion.Text);
            Console.WriteLine(criterionAnswerTemplate);
            Console.WriteLine(OutputStrings.FinishOption);
            var input = Console.ReadKey();
            criterionAnswer = input.Key switch
            {
                ConsoleKey.D1 => OutputStrings.NoDifficulties,
                ConsoleKey.D2 => OutputStrings.SmallDifficulties,
                ConsoleKey.D3 => OutputStrings.MiddleDifficulties,
                ConsoleKey.D4 => OutputStrings.HardDifficulties,
                ConsoleKey.D5 => OutputStrings.TotalDifficulties,
                ConsoleKey.Escape => throw new FinishFillingException(),
                _ => null
            };
            if (criterionAnswer is null)
            {
                Console.WriteLine(OutputStrings.WrongInputMessage);
            }
        } while (string.IsNullOrEmpty(criterionAnswer));

        return criterionAnswer;
    }
    
    private static void AddCriterionRow(Table table, Criterion criterion, string criterionAnswer)
    {
        table.AppendChild(new TableRow(
            new TableCell(
                new TableCellProperties(
                    new TableCellWidth {Width = DocumentMetrics.TableCellWidth}),
                new Paragraph(
                    new ParagraphProperties(
                        new Justification {Val = new EnumValue<JustificationValues>(JustificationValues.Left)}),
                    new Run(
                        new RunProperties(
                            new FontSize {Val = new StringValue(DocumentMetrics.Fonts.CriterionFontSize)},
                            new FontSizeComplexScript {Val = new StringValue(DocumentMetrics.Fonts.CriterionFontSize)}),
                        new Text(criterion.Text)))),
            new TableCell(
                new TableCellProperties(
                    new TableCellWidth {Width = DocumentMetrics.TableCellWidth}),
                new Paragraph(
                    new ParagraphProperties(
                        new Justification {Val = new EnumValue<JustificationValues>(JustificationValues.Left)}),
                    new Run(
                        new RunProperties(
                            new FontSize {Val = new StringValue(DocumentMetrics.Fonts.CriterionFontSize)},
                            new FontSizeComplexScript {Val = new StringValue(DocumentMetrics.Fonts.CriterionFontSize)}),
                        new Text(criterionAnswer))))));
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