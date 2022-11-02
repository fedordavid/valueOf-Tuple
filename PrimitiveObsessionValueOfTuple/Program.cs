using FluentValidation;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using PrimitiveObsessionValueOfTuple.ExcelFileReader;
using PrimitiveObsessionValueOfTuple.Model;
using PrimitiveObsessionValueOfTuple.Validator;

const string Path = @"C:\Users\ZZ020T693\Desktop\Desktop\FIFA\UserStories\19161 Individual Bulk Import - Bug\Test.xlsx";

// Open File From Disk

// Dependency Injection
using var host = Host.CreateDefaultBuilder(args)
    .ConfigureServices((_, services) =>
        services.AddValidatorsFromAssemblyContaining<IndividualImportRowValidator>())
    .Build();

// Get Required Services
using var serviceScope = host.Services.CreateScope();
var provider = serviceScope.ServiceProvider;
var validator = provider.GetRequiredService<IValidator<IndividualImportRow>>();

var stream = File.OpenRead(Path);

ExcelFileReader<IndividualImportRow> CreateIndividualImportRowReader()
{
    return new ExcelFileReader<IndividualImportRow>()
        .AddColumn((individual, cell) => individual.FirstName = cell.Text)
        .AddColumn((individual, cell) => individual.LastName = cell.Text)
        .AddColumn((individual, cell) => individual.Gender = cell.Text)
        .AddColumn((individual, cell) => individual.FifaNationality = cell.Text)
        .AddColumn((individual, cell) => individual.BirthDate = cell.Text.GetCellDate(cell))
        .AddColumn((individual, cell) => individual.CorrespondenceLanguage = cell.Text)
        .AddColumn((individual, cell) => individual.Email = cell.Text)
        .AddColumn((individual, cell) => individual.IdentificationType = cell.Text)
        .AddColumn((individual, cell) => individual.IdentificationCountry = cell.Text)
        .AddColumn((individual, cell) => individual.IdentificationDocumentNumber = cell.Text)
        .AddColumn((individual, cell) => individual.IdentificationPlaceOfIssue = cell.Text)
        .AddColumn((individual, cell) => individual.IdentificationCountryOfIssue = cell.Text)
        .AddColumn((individual, cell) => individual.IdentificationAuthority = cell.Text)
        .AddColumn((individual, cell) =>
        {
            individual.NewIdentificationDateOfIssue = NewIdentificationDateOfIssue.From(cell.Text.TryConvertToDateTime("dd.MM.yyyy"));
                // .AddColumn((individual, cell) => individual.IdentificationDateOfIssue = cell.Text.GetCellDate(cell))
        })
        .AddColumn((individual, cell) => individual.IdentificationExpirationDate = cell.Text.GetCellDate(cell));
}

var excelReader = CreateIndividualImportRowReader()
    .ReadFromStream(stream)
    .Select(individual => (individual, errors: validator.Validate(individual).Errors.Select(v => v.ErrorMessage).ToList())).ToList();

static bool IsValidationSuccessful(IEnumerable<(IndividualImportRow individual, List<string> errors)> individualImportRows)
    => individualImportRows.All(row => row.errors.Count == 0);

Console.WriteLine(IsValidationSuccessful);

