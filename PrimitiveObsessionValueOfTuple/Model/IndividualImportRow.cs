namespace PrimitiveObsessionValueOfTuple.Model;

public class IndividualImportRow
{
    public string FirstName { get; set; }
        
    public string LastName { get; set; }
            
    public string Gender { get; set; }
            
    public string FifaNationality { get; set; }
            
    public DateTime? BirthDate { get; set; }
            
    public string CorrespondenceLanguage { get; set; }
            
    public string Email { get; set; }
            
    public string IdentificationType { get; set; }
            
    public string IdentificationCountry { get; set; }
            
    public string IdentificationDocumentNumber { get; set; }
            
    public string IdentificationPlaceOfIssue { get; set; }
            
    public string IdentificationCountryOfIssue { get; set; }
            
    public string IdentificationAuthority { get; set; }
            
    public DateTime? IdentificationDateOfIssue { get; set; }
            
    public DateTime? IdentificationExpirationDate { get; set; }
    
    public NewIdentificationDateOfIssue NewIdentificationDateOfIssue { get; set; }
        
    public List<string> ValidationError { get; set; }
    
}