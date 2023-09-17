using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ExistingReportFiller;
using ExistingReportFiller.Models;
using Shared;

internal class Program
{
    public static void Main(string[] args)
    {
        var unitsFilePath = @"C:\Users\merkulov.e\Source\Playground\ReportGenerator\CriteriaParser\JsonView.json";
        string existingFilePath = null;
        if (args.Any())
        {
            foreach (var argument in args)
            {
                switch (argument)
                {
                    case "unitsFilePath":
                        unitsFilePath = argument;
                        break;
                    case "existingFilePath":
                        existingFilePath = argument;
                        break;
                }
            }
        }

        if (string.IsNullOrEmpty(existingFilePath))
        {
            Console.WriteLine("Не указан путь к заполняемому файлу");
            return;
        }
        
        using var existingDocument = WordprocessingDocument.Open(existingFilePath, true);
        var criteriaTable = GetReportTable(existingDocument);

        if (criteriaTable is null)
        {
            Console.WriteLine("В документе не найдена таблица с критериями. Если вы еще не начинали ее заполнять, создайте новый документ.");
        }

        var filledCriteriaIds = CollectFilledCriteriaIds(criteriaTable);
        Console.WriteLine("Документ обработан");

        do
        {
            var action = SelectAction();
            switch (action)
            {
                case UserAction.InputNewCriteria:
                    InputNewCriteria(filledCriteriaIds, criteriaTable);
                    break;
                case UserAction.ChangeOldCriteria:
                    break;
                case UserAction.DeleteOldCriteria:
                    break;
                case UserAction.Finish:
                    return;
            }
        } while (true);

    }
    
    private static Table GetReportTable(WordprocessingDocument existingDocument)
    {
        var body = existingDocument.MainDocumentPart!.Document.Body;
        return body!.ChildElements
            .OfType<Table>()
            .FirstOrDefault();
    }
    
    private static HashSet<int> CollectFilledCriteriaIds(Table criteriaTable)
    {
        var rowProperties = criteriaTable!.ChildElements
            .OfType<SdtRow>()
            .SelectMany(row => row.ChildElements.OfType<SdtProperties>());
        
        return rowProperties
            .Where(props => props.Any(prop => 
                prop.ChildElements.OfType<SdtAlias>()
                .Any(alias => alias!.Val!.Value == nameof(Criterion))))
            .SelectMany(properties => properties.OfType<SdtId>())
            .Select(id => id!.Val!.Value)
            .ToHashSet();
    }
    
    private static UserAction SelectAction()
    {
        var selectActionTemplate = 
            $"""
                 Выберите что вы хотите сделать:
                     {(int)UserAction.InputNewCriteria} - {OutputStrings.UserActionsMessages.InputNewCriteriaUserActionMessage}
                     {(int)UserAction.ChangeOldCriteria} - {OutputStrings.UserActionsMessages.ChangeOldCriteriaUserActionMessage}
                     {(int)UserAction.DeleteOldCriteria} - {OutputStrings.UserActionsMessages.DeleteOldCriteriaUserActionMessage}
                     Esc - {OutputStrings.UserActionsMessages.FinishActionMessage}
             """;

        UserAction action;
        do
        {
            Console.WriteLine(selectActionTemplate);
            var input = Console.ReadKey();
            action = input.Key switch
            {
                ConsoleKey.D1 => UserAction.InputNewCriteria,
                ConsoleKey.D2 => UserAction.ChangeOldCriteria,
                ConsoleKey.D3 => UserAction.DeleteOldCriteria,
                ConsoleKey.Escape => UserAction.Finish,
                _ => UserAction.Undefined
            };

            if (action == UserAction.Undefined)
            {
                Console.WriteLine(OutputStrings.WrongInputMessage);
            }
            
        } while (action == UserAction.Undefined);

        return action;
    }
    
    private static void InputNewCriteria(HashSet<int> filledCriteriaIds, Table table)
    {
        var criteriaId = InputCriteriaId();

        if (filledCriteriaIds.Contains(criteriaId))
        {
            Console.WriteLine(OutputStrings.NewCriteriaMessages.CriteriaAlreadyExistsMessage);
            return;
        }

        var criteriaUnit = FindCriteriaUnitRow(criteriaId, table);
    }
    
    private static int InputCriteriaId()
    {
        int criteriaId;
        do
        {
            Console.WriteLine(OutputStrings.InputCriteriaMessages.InputCriteriaMessage);
            if (!int.TryParse(Console.ReadLine(), out criteriaId))
            {
                Console.WriteLine(OutputStrings.InputCriteriaMessages.InvalidCriteriaIdMessage);
            }
            
        } while (criteriaId == 0);

        return criteriaId;
    }
    
    private static SdtRow FindCriteriaUnitRow(int criteriaId, Table table)
    {
        var rowProperties = table!.ChildElements
            .OfType<SdtRow>()
            .FirstOrDefault(row => row.Where(row => 
                row.ChildElements.OfType<SdtProperties>().Any(props => 
                    props.Any(prop => prop.ChildElements.OfType<SdtAlias>()))))
            .SelectMany(row => row.OfType<SdtProperties>());
        
        return rowProperties
            .Where(props => props.Any(prop => 
                prop.ChildElements.OfType<SdtAlias>()
                    .Any(alias => alias!.Val!.Value == nameof(Criterion))))
            .SelectMany(properties => properties.OfType<SdtId>())
            .Select(id => id!.Val!.Value)
            .ToHashSet();
    }

}