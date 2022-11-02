using FluentValidation;
using PrimitiveObsessionValueOfTuple.Model;

namespace PrimitiveObsessionValueOfTuple.Validator;

public class NewIdentificationDateOfIssueValidator : AbstractValidator<NewIdentificationDateOfIssue>
{
    public NewIdentificationDateOfIssueValidator()
    {
        RuleFor(e => e.Value.convertSuccess)
            .NotEqual(false)
            .When(HasValueProvided)
            .WithMessage("Not a valid identification date of issue provided.");
    }

    private static bool HasValueProvided(NewIdentificationDateOfIssue arg) 
        => !string.IsNullOrEmpty(arg.Value.stringValueOfDateTime) && arg.Value.dateTime is null;
}