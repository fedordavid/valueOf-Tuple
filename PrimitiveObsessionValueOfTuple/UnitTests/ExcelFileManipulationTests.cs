using FluentValidation;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PrimitiveObsessionValueOfTuple.ExcelFileReader;
using PrimitiveObsessionValueOfTuple.Model;
using PrimitiveObsessionValueOfTuple.Validator;

namespace PrimitiveObsessionValueOfTuple.Tests;

[TestClass]
public class ExcelFileManipulationTests
{
    private const string Path =
        @"C:\Users\ZZ020T693\Desktop\Desktop\FIFA\UserStories\19161 Individual Bulk Import - Bug\Test.xlsx";

    private const string PathError =
        @"C:\Users\ZZ020T693\Desktop\Desktop\FIFA\UserStories\19161 Individual Bulk Import - Bug\TestError.xlsx";

    private IServiceProvider _provider;
    private IValidator<IndividualImportRow> _validator;

    [TestInitialize]
    public void TestInitialize()
    {
        using var host = Host.CreateDefaultBuilder()
            .ConfigureServices((_, services) =>
                services.AddValidatorsFromAssemblyContaining<IndividualImportRowValidator>())
            .Build();

        using var serviceScope = host.Services.CreateScope();
        _provider = serviceScope.ServiceProvider;
        _validator = _provider.GetRequiredService<IValidator<IndividualImportRow>>();
    }

    [TestMethod]
    [TestCategory("ProfileAPI")]
    public void SerializeIndividualImportFile()
    {
        var individualImportBodyDefinitions = new List<IndividualImportRow>()
        {
            new()
            {
                FirstName = "Ezechiel",
                LastName = "Dufour",
                Gender = "Female",
                BirthDate = DateTime.Now,
                FifaNationality = "France",
                CorrespondenceLanguage = "English",
                Email = "david@fedor.sk",
                IdentificationType = "Passport",
                IdentificationCountry = "Albania",
                IdentificationDocumentNumber = "1234567899",
                IdentificationPlaceOfIssue = "whatever",
                IdentificationCountryOfIssue = "Albania",
                IdentificationAuthority = "police",
                IdentificationDateOfIssue = DateTime.Now,
                IdentificationExpirationDate = DateTime.Today.AddYears(10)
            }
        };

        var tuple = individualImportBodyDefinitions
            .Select(individualImportRow => (individualImportRow, errors: _validator
                .Validate(individualImportRow)
                .Errors.Select(v => v.ErrorMessage)
                .ToList())).ToList();

        using var stream = File.OpenRead(Path);

        new ExcelFileWriter<IndividualImportRow>()
            .AddColumn((individual, cell, _) => cell.Value = individual.FirstName)
            .AddColumn((individual, cell, _) => cell.Value = individual.LastName)
            .AddColumn((individual, cell, _) => cell.Value = individual.Gender)
            .AddColumn((individual, cell, _) => cell.Value = individual.FifaNationality)
            .AddColumn((_, cell, _) => cell.Value = cell.Value)
            .AddColumn((individual, cell, _) => cell.Value = individual.CorrespondenceLanguage)
            .AddColumn((individual, cell, _) => cell.Value = individual.Email)
            .AddColumn((individual, cell, _) => cell.Value = individual.IdentificationType)
            .AddColumn((individual, cell, _) => cell.Value = individual.IdentificationCountry)
            .AddColumn((individual, cell, _) => cell.Value = individual.IdentificationDocumentNumber)
            .AddColumn((individual, cell, _) => cell.Value = individual.IdentificationPlaceOfIssue)
            .AddColumn((individual, cell, _) => cell.Value = individual.IdentificationCountryOfIssue)
            .AddColumn((individual, cell, _) => cell.Value = individual.IdentificationAuthority)
            .AddColumn((_, cell, _) => cell.Value = cell.Value)
            .AddColumn((_, cell, _) => cell.Value = cell.Value)
            .AddColumn((_, cell, errors) => cell.RichText.AddRichTextError(errors))
            .WriteToStream(tuple, stream);
    }

    [TestMethod]
    [TestCategory("ProfileAPI")]
    public void DeserializeIndividualImportFile()
    {
        using var stream = File.OpenRead(Path);
        var individualImportBodyDefinitions = new ExcelFileReader<IndividualImportRow>()
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
                individual.NewIdentificationDateOfIssue =
                    NewIdentificationDateOfIssue.From(cell.Text.TryConvertToDateTime("dd.MM.yyyy"));
            })
            .AddColumn((individual, cell) => individual.IdentificationExpirationDate = cell.Text.GetCellDate(cell))
            .ReadFromStream(stream)
            .Select(item => (item, _validator.Validate(item).Errors.Select(v => v.ErrorMessage).ToList()))
            .ToArray();

        if (IsValidationSuccessful(individualImportBodyDefinitions)) return;


        var excel = new ExcelFileWriter<IndividualImportRow>()
            .AddColumn((individual, cell, _) => cell.Value = individual.FirstName)
            .AddColumn((individual, cell, _) => cell.Value = individual.LastName)
            .AddColumn((individual, cell, _) => cell.Value = individual.Gender)
            .AddColumn((individual, cell, _) => cell.Value = individual.FifaNationality)
            .AddColumn((_, cell, _) => cell.Value = cell.Value)
            .AddColumn((individual, cell, _) => cell.Value = individual.CorrespondenceLanguage)
            .AddColumn((individual, cell, _) => cell.Value = individual.Email)
            .AddColumn((individual, cell, _) => cell.Value = individual.IdentificationType)
            .AddColumn((individual, cell, _) => cell.Value = individual.IdentificationCountry)
            .AddColumn((individual, cell, _) => cell.Value = individual.IdentificationDocumentNumber)
            .AddColumn((individual, cell, _) => cell.Value = individual.IdentificationPlaceOfIssue)
            .AddColumn((individual, cell, _) => cell.Value = individual.IdentificationCountryOfIssue)
            .AddColumn((individual, cell, _) => cell.Value = individual.IdentificationAuthority)
            .AddColumn((_, cell, _) => cell.Value = cell.Value)
            .AddColumn((_, cell, _) => cell.Value = cell.Value)
            .AddColumn((_, cell, errors) =>
            {
                cell.RichText.Clear();
                cell.RichText.AddRichTextError(errors);
            })
            .WriteToStream(individualImportBodyDefinitions, stream);

        File.WriteAllBytes(PathError, excel);


        static bool IsValidationSuccessful(
            IEnumerable<(IndividualImportRow individual, List<string> errors)> individualImportRows)
            => individualImportRows.All(row => row.errors.Count == 0);
    }
}