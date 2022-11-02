using OfficeOpenXml;
using PrimitiveObsessionValueOfTuple.Model;

namespace PrimitiveObsessionValueOfTuple.ExcelFileReader;

public class ExcelFileWriter<TItem> where TItem : class
{
    private readonly List<Action<TItem, ExcelRange, IReadOnlyCollection<string>>> _cells = new();
    private readonly List<Action<ExcelColumn>> _formats = new();

    public ExcelFileWriter<TItem> AddColumn(Action<TItem, ExcelRange, IReadOnlyCollection<string>> parseCell, Action<ExcelColumn> configureColumnStyle = null)
    {
        _cells.Add(parseCell);
        _formats.Add(configureColumnStyle ?? (_ => {}));
        return this;
    }

    public byte[] WriteToStream(IEnumerable<(IndividualImportRow individualImportRow, List<string> errors)> importRows, Stream stream)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using var package = new ExcelPackage();
        package.Load(stream);
            
        var sheet = package.Workbook.Worksheets.First();
        var rows = importRows.ToArray();
           
        for (var row = 0; row < rows.Length; row++)
        {
            WriteRow(sheet, row + 1, rows[row].individualImportRow as TItem, rows[row].errors);
        }

        for (var col = 0; col < _formats.Count(); col++)
        {
            _formats[col](sheet.Column(col + 1));
        }
            
        return package.GetAsByteArray();
    }

    private void WriteRow(ExcelWorksheet sheet, int row, TItem item, IReadOnlyCollection<string> errors)
    {
        for (var col = 0; col < _cells.Count; col++)
        {
            var cell = _cells[col];
            cell(item, sheet.Cells[row + 1, col + 1], errors);
        }
    }
}