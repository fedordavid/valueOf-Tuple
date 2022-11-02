using System.Text.RegularExpressions;
using FluentValidation;
using PrimitiveObsessionValueOfTuple.Model;

namespace PrimitiveObsessionValueOfTuple.Validator;

public class IndividualImportRowValidator : AbstractValidator<IndividualImportRow>
    {
        private const string EmailRegex = @"^[a-zA-Z0-9]([a-zA-Z0-9_\.\-]*)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|([a-zA-Z0-9]([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$";

        public IndividualImportRowValidator()
        {
            RuleFor(e => e.NewIdentificationDateOfIssue).SetValidator(new NewIdentificationDateOfIssueValidator());
                            
            RuleFor(e => e.FirstName)
                .Length(1, 30)
                .WithMessage("First name must be between 1 and 30 characters.");
            
            RuleFor(e => e.LastName)
                .Length(1,30)
                .WithMessage("Last name must be between 1 and 30 characters.");
            
            RuleFor(e => e)
                .Must(FullNameLength)
                .WithMessage("First name and last name must be less than 50 characters.")
                .WithName("First name and last name");

            RuleFor(e => e.Gender)
                .Must(HaveValidGender)
                .WithMessage("Not a valid gender provided.");
            
            RuleFor(e => e.BirthDate)
                .NotNull()
                .WithMessage("Not a valid birthdate provided.");

            RuleFor(e => e.Email)
                .Length(0, 60)
                .Must(BeAValidAddress)
                .When(x => !string.IsNullOrEmpty(x.Email));

            RuleFor(e => e.IdentificationDocumentNumber)
                .Length(1,60)
                .When(AtLeastOneIdentificationFieldProvided)
                .WithMessage("Not a valid identification document number provided.");

            RuleFor(e => e.IdentificationPlaceOfIssue)
                .Length(0,255);
            
            RuleFor(e => e.IdentificationCountryOfIssue)
                .Length(0, 255)
                .When(AtLeastOneIdentificationFieldProvided)
                .WithMessage("Not a valid identification country of issue provided.");

            RuleFor(e => e.IdentificationAuthority)
                .Length(0, 255)
                .WithMessage("Identification authority must be between 1 and 255 characters.");
            
            RuleFor(e => e.IdentificationDateOfIssue)
                .Must(BeInThePast)
                .When(e => e.IdentificationDateOfIssue is not null)
                .WithMessage("Identification date of issue has to be in the past.");
            
            RuleFor(e => e.IdentificationExpirationDate)
                .Must(BeInTheFuture)
                .When(AtLeastOneIdentificationFieldProvided)
                .WithMessage("Identification date of expiration has to be in the future.");
        }
        
        private static bool BeAValidAddress(string emailAddress) 
            => Regex.IsMatch(emailAddress, EmailRegex, RegexOptions.IgnoreCase);

        private static bool AtLeastOneIdentificationFieldProvided(IndividualImportRow individualImportRow) =>
            !string.IsNullOrEmpty(individualImportRow.IdentificationType) ||
            !string.IsNullOrEmpty(individualImportRow.IdentificationCountry) ||
            !string.IsNullOrEmpty(individualImportRow.IdentificationDocumentNumber) ||
            !string.IsNullOrEmpty(individualImportRow.IdentificationCountryOfIssue) ||
            individualImportRow.IdentificationExpirationDate != null;

        private static bool HaveValidGender(string gender) => gender is "Male" or "Female";

        private static bool BeInThePast(DateTime? date) => date != null && date < DateTime.Now;

        private static bool BeInTheFuture(DateTime? date) => date != null && date > DateTime.Now;

        private static bool FullNameLength(IndividualImportRow individualImportRow)
        {
            var firstName = individualImportRow.FirstName?.Length ?? 0;
            var lastName = individualImportRow.LastName?.Length ?? 0;
            return firstName + lastName <= 50;
        } 
    }