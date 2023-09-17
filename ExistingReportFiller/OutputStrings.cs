namespace ExistingReportFiller;

public static class OutputStrings
{
    public const string WrongInputMessage = "Выбран неподходящий вариант ответа";
    
    public static class UserActionsMessages
    {
        public const string InputNewCriteriaUserActionMessage = "Ввести новый критерий";
        public const string ChangeOldCriteriaUserActionMessage = "Поменять значение старого критерия";
        public const string DeleteOldCriteriaUserActionMessage = "Удалить старый критерия";
        public const string FinishActionMessage = "Сохранить документ и закончить работу.";
    }
    
    public static class InputCriteriaMessages
    {
        public const string InputCriteriaMessage = "Введите номер критерия";
        public const string InvalidCriteriaIdMessage = "Некоректный номер критерия. Попробуйте еще раз";
    }
    
    public static class NewCriteriaMessages
    {
        public const string CriteriaAlreadyExistsMessage =
            "Данные критерий уже существует в документе. Если вы хотите заменить его значение выберите соотсветсвующую опцию.";
    }
}