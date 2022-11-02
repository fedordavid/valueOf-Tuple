using System.Globalization;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace PrimitiveObsessionValueOfTuple.ExcelFileReader;

public static class ExcelRangeParsingExtensions
{
    public static DateTime? GetCellDate(this string str, ExcelRange range) =>
        range.Text.ToNullableDateTime("dd.MM.yyyy");

    public static void AddRichTextError(this ExcelRichTextCollection richTextCollection,
        IReadOnlyCollection<string> errors)
    {
        if (errors is null)
        {
            richTextCollection.Add(string.Empty);
            return;
        }

        var index = 0;
        foreach (var error in errors)
        {
            index++;
            if (index == errors.Count)
            {
                richTextCollection.Add(error);
            }
            else
            {
                richTextCollection.Add(error + "\r\n");
            }
        }
    }

    private static DateTime? ToNullableDateTime(this string value, string format)
    {
        DateTime? result = null;
        DateTime time;
        if (DateTime.TryParseExact(value, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out time))
        {
            result = time;
        }

        return result;
    }

    public static (bool convertSuccess, DateTime? dateTime, string stringValueOfDateTime) TryConvertToDateTime(this string value, string format)
    {
        var success = false;
        DateTime? result = null;

        if (DateTime.TryParseExact(value, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out var time))
        {
            result = time;
            success = true;
        }

        return (success, result, value);
    }
}