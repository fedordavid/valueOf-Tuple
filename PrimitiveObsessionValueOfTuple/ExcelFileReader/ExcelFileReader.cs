using OfficeOpenXml;

namespace PrimitiveObsessionValueOfTuple.ExcelFileReader;

public class ExcelFileReader<TItem> where TItem : new()
{
    private readonly List<Action<TItem, ExcelRange>> _cells = new();

    public ExcelFileReader<TItem> AddColumn(Action<TItem, ExcelRange> parseCell)
    {
        _cells.Add(parseCell);
        return this;
    }

    public IEnumerable<TItem> ReadFromStream(Stream stream)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using var package = new ExcelPackage();
        package.Load(stream);

        if (package.Workbook.Worksheets.Count == 0)
        {
            throw new ApplicationException("Could not load Excel Sheet");
        }
            
        var sheet = package.Workbook.Worksheets.First();
        var rows = Rows(sheet);

        foreach (var row in rows)
        {
            if (!RowIsEmpty(sheet, row))
            {
                yield return ParseRow(sheet, row, new TItem());     
            }
        }
    }

    private TItem ParseRow(ExcelWorksheet sheet, int row, TItem item)
    {
        for (var col = 0; col < _cells.Count; col++)
        {
            var cell = _cells[col];
            cell(item, sheet.Cells[row, col + 1]);
        }

        return item;
    }
        
    private static IEnumerable<int> Rows(ExcelWorksheet worksheet) 
        => Enumerable.Range(2, worksheet.Dimension.Rows - 1);
        
    private static IEnumerable<int> Cols(ExcelWorksheet worksheet) 
        => Enumerable.Range(1, worksheet.Dimension.End.Column);
        
    private static bool RowIsEmpty(ExcelWorksheet xlsWorksheet, int row) 
        => Cols(xlsWorksheet).All(col => xlsWorksheet.Cells[row, col].Value == null);
}