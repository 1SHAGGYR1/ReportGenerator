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
        AddChildInfo(body, info);
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

    private static void AddChildInfo(Body body, DocumentInfo info)
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
                    new Text($"{string.Concat(info.ChildInfo.LastName[0].ToString().ToUpper(), info.ChildInfo.LastName.AsSpan(1))} {string.Concat(info.ChildInfo.FirstName[0].ToString().ToUpper(), info.ChildInfo.FirstName.AsSpan(1))}: {info.PeriodStartDate} - {info.PeriodEndDate}"))
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
                new GridColumn { Width = new StringValue(DocumentMetrics.TableCellWidth) },
                new GridColumn { Width = new StringValue(DocumentMetrics.TableCellWidth) }),
            new TableRow(
                new TableCell(
                    new TableCellProperties(
                        new TableCellWidth { Width = DocumentMetrics.TableWidth },
                        new GridSpan{Val = new Int32Value(3)}),
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
                if (!InputPart("раздела", unit.Text)) continue;
                AddUnitRow(table, unit);

                foreach (var section in unit.SectionList)
                {
                    if (!InputPart("секции", section.Text)) continue;
                    AddSectionRow(table, section);

                    foreach (var criterion in section.CriterionList)
                    {
                        ProcessCriterion(table, criterion);
                    }
                }

                foreach (var unSectionedCriterion in unit.UnSectionedCriterionList)
                {
                    ProcessCriterion(table, unSectionedCriterion);
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
    
    private static void AddUnitRow(Table table, Unit unit)
    {
        table.AppendChild(
            new SdtRow(
                new SdtProperties(
                    new SdtAlias {Val = new StringValue(nameof(Unit))},
                    new SdtId {Val = new Int32Value(unit.Digit)}),
                new SdtContentRow(
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(
                                new TableCellWidth {Width = DocumentMetrics.TableWidth},
                                new GridSpan {Val = new Int32Value(3)}),
                            new Paragraph(
                                new ParagraphProperties(
                                    new Justification
                                        {Val = new EnumValue<JustificationValues>(JustificationValues.Center)}),
                                new Run(
                                    new RunProperties(
                                        new Bold(),
                                        new FontSize {Val = new StringValue(DocumentMetrics.Fonts.UnitFontSize)},
                                        new FontSizeComplexScript
                                            {Val = new StringValue(DocumentMetrics.Fonts.UnitFontSize)}),
                                    new Text(unit.Text))))))));
    }

    private static void AddSectionRow(Table table, Section section)
    {
        table.AppendChild(
            new SdtRow(
                new SdtProperties(
                    new SdtAlias {Val = new StringValue(nameof(Section))},
                    new SdtId {Val = new Int32Value(section.StartCriterionKey)}),
                new SdtContentRow(
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(
                                new TableCellWidth {Width = DocumentMetrics.TableWidth},
                                new GridSpan {Val = new Int32Value(3)}),
                            new Paragraph(
                                new ParagraphProperties(
                                    new Justification
                                        {Val = new EnumValue<JustificationValues>(JustificationValues.Center)}),
                                new Run(
                                    new RunProperties(
                                        new Italic(),
                                        new FontSize {Val = new StringValue(DocumentMetrics.Fonts.SectionFontSize)},
                                        new FontSizeComplexScript
                                            {Val = new StringValue(DocumentMetrics.Fonts.SectionFontSize)}),
                                    new Text(section.Text))))))
            ));
    }
    
    private static void ProcessCriterion(Table table, Criterion criterion)
    {
        var criterionAnswer = InputCriterionAnswer(criterion);
        if (criterionAnswer == CriterionAnswers.Skip)
        {
            return;
        }
        
        AddCriterionRow(table, criterion, criterionAnswer);

        if (criterion.Inner is not null)
        {
            foreach (var innerCriterion in criterion.Inner.Values)
            {
                ProcessCriterion(table, innerCriterion);
            }
        }
    }
    
    private static bool InputPart(string partName, string partText)
    {
        Console.WriteLine(OutputStrings.StartFillingPartMessage, partName, partText);
        bool? inputUnit;
        do
        {
            Console.WriteLine(OutputStrings.SkipOption, partName);
            Console.WriteLine(OutputStrings.ContinueOption, partName);
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

    private static CriterionAnswers InputCriterionAnswer(Criterion criterion)
    {
        var criterionAnswerTemplate =
            $"""
            Для заполнения критерия введите уровень затруднений:
                {(int)CriterionAnswers.NoDifficulties} - {OutputStrings.NoDifficulties}
                {(int)CriterionAnswers.SmallDifficulties} - {OutputStrings.SmallDifficulties}
                {(int)CriterionAnswers.MiddleDifficulties} - {OutputStrings.MiddleDifficulties}
                {(int)CriterionAnswers.HardDifficulties} - {OutputStrings.HardDifficulties}
                {(int)CriterionAnswers.TotalDifficulties} - {OutputStrings.TotalDifficulties}
            """;
        CriterionAnswers criterionAnswer;
        do
        {
            Console.WriteLine(OutputStrings.StartFillingPartMessage, "критерия", criterion.Text);
            Console.WriteLine(OutputStrings.SkipOption, "критерия");
            Console.WriteLine(OutputStrings.FinishOption);
            Console.WriteLine(criterionAnswerTemplate);
            var input = Console.ReadKey();
            criterionAnswer = input.Key switch
            {
                ConsoleKey.D0 => CriterionAnswers.Skip,
                ConsoleKey.D1 => CriterionAnswers.NoDifficulties,
                ConsoleKey.D2 => CriterionAnswers.SmallDifficulties,
                ConsoleKey.D3 => CriterionAnswers.MiddleDifficulties,
                ConsoleKey.D4 => CriterionAnswers.HardDifficulties,
                ConsoleKey.D5 => CriterionAnswers.TotalDifficulties,
                ConsoleKey.Escape => throw new FinishFillingException(),
                _ => CriterionAnswers.Undefined
            };
            if (criterionAnswer is CriterionAnswers.Undefined)
            {
                Console.WriteLine(OutputStrings.WrongInputMessage);
            }
        } while (criterionAnswer == CriterionAnswers.Undefined);

        Console.WriteLine();

        return criterionAnswer;
    }

    private static void AddCriterionRow(Table table, Criterion criterion, CriterionAnswers criterionAnswer)
    {
        var cellProperties = new TableCellProperties(new TableCellWidth {Width = DocumentMetrics.TableCellWidth});
        if (criterionAnswer != CriterionAnswers.NoDifficulties)
        {
            cellProperties.AppendChild(new Shading
            {
                Color = new StringValue("auto"),
                Val = new EnumValue<ShadingPatternValues>(ShadingPatternValues.Clear),
                Fill = new StringValue(ChooseShadingColor(criterionAnswer))
            });
        }

        table.AppendChild(
            new SdtRow(
                new SdtProperties(
                    new SdtAlias {Val = new StringValue(nameof(Criterion))},
                    new SdtId {Val = new Int32Value(criterion.Key)}),
                new SdtContentRow(new TableRow(
                    new TableCell(
                        cellProperties,
                        new Paragraph(
                            new ParagraphProperties(
                                new Justification {Val = new EnumValue<JustificationValues>(JustificationValues.Left)}),
                            new Run(
                                new RunProperties(
                                    new FontSize {Val = new StringValue(DocumentMetrics.Fonts.CriterionFontSize)},
                                    new FontSizeComplexScript
                                        {Val = new StringValue(DocumentMetrics.Fonts.CriterionFontSize)}),
                                new Text(criterion.Text)))),
                    new TableCell(
                        cellProperties.CloneNode(true),
                        new Paragraph(
                            new ParagraphProperties(
                                new Justification {Val = new EnumValue<JustificationValues>(JustificationValues.Left)}),
                            new Run(
                                new RunProperties(
                                    new FontSize {Val = new StringValue(DocumentMetrics.Fonts.CriterionFontSize)},
                                    new FontSizeComplexScript
                                        {Val = new StringValue(DocumentMetrics.Fonts.CriterionFontSize)}),
                                new Text(ChooseAnswerText(criterionAnswer))))),
                    new TableCell(
                        cellProperties.CloneNode(true),
                        new Paragraph(
                            new ParagraphProperties(
                                new Justification {Val = new EnumValue<JustificationValues>(JustificationValues.Left)}),
                            new Run(
                                new RunProperties(
                                    new FontSize {Val = new StringValue(DocumentMetrics.Fonts.CriterionFontSize)},
                                    new FontSizeComplexScript
                                        {Val = new StringValue(DocumentMetrics.Fonts.CriterionFontSize)}))))))
            ));

        string ChooseAnswerText(CriterionAnswers answer) => answer switch
        {
            CriterionAnswers.NoDifficulties => OutputStrings.NoDifficulties,
            CriterionAnswers.SmallDifficulties => OutputStrings.SmallDifficulties,
            CriterionAnswers.MiddleDifficulties => OutputStrings.MiddleDifficulties,
            CriterionAnswers.HardDifficulties => OutputStrings.HardDifficulties,
            CriterionAnswers.TotalDifficulties => OutputStrings.TotalDifficulties,
            _ => throw new ArgumentOutOfRangeException(nameof(answer), answer, null)
        };

        string ChooseShadingColor(CriterionAnswers answer) => answer switch
        {
            CriterionAnswers.NoDifficulties => DocumentMetrics.TableShadingColors.Green,
            CriterionAnswers.SmallDifficulties => DocumentMetrics.TableShadingColors.Green,
            CriterionAnswers.MiddleDifficulties => DocumentMetrics.TableShadingColors.Yellow,
            CriterionAnswers.HardDifficulties => DocumentMetrics.TableShadingColors.Orange,
            CriterionAnswers.TotalDifficulties => DocumentMetrics.TableShadingColors.Red,
            _ => throw new ArgumentOutOfRangeException(nameof(answer), answer, null)
        };
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