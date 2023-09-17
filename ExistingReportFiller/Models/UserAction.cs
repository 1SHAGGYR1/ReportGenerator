namespace ExistingReportFiller.Models;

public enum UserAction
{
    Undefined = 0,
    InputNewCriteria = 1,
    ChangeOldCriteria = 2,
    DeleteOldCriteria = 3,
    Finish = 10
}