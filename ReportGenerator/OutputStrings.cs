namespace ReportGenerator;

public static class OutputStrings
{
    public const string NoDifficulties = "НЕТ затруднений (никаких, отсутствуют, незначительные,…) 0-4%";
    public const string SmallDifficulties = "ЛЕГКИЕ затруднения (незначительные, слабые,…) 5-24%";
    public const string MiddleDifficulties = "УМЕРЕННЫЕ затруднения (средние, значимые,…) 25-49%";
    public const string HardDifficulties = "ТЯЖЕЛЫЕ затруднения (высокие, интенсивные,…) 50-95%";
    public const string TotalDifficulties = "АБСОЛЮТНЫЕ затруднения (полные,…) 96-100%";

    public const string WrongInputMessage = "Выбран неподходящий вариант ответа";
    public const string StartFillingPartMessage = "Переходим к заполнению {0}: {1}";

    public const string SkipOption = "Для пропуска {0} нажмите 0, для ввода: ENTER";
    public const string FinishOption = "Для завершения заполнения нажмите ESC";
}