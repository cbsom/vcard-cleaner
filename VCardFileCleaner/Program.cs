using System.Text;
using System.Text.RegularExpressions;
using MimeKit.Encodings;
using Sylvan.Data;
using Sylvan.Data.Csv;

/**
 * This program loads either a vCard (.vcf) file or a csv file from the supplied path, 
 * parses each contact into objects, 
 * cleans the phone numbers if necessary, 
 * merges the records that share the same name but have different phone numbers, 
 * removes duplicates, 
 * writes the distinct records back to a new vcf file and/or a new csv file.
 */
partial class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.WriteLine("Usage: VCardFileCleaner.exe <input file path>");
            DoExit();
            return;
        }
        var path = args[0];
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            Console.WriteLine("Invalid path: " + path);
            DoExit();
            return;
        }
        var records = Path.GetExtension(path).ToLowerInvariant() == ".vcf"
            ? LoadVCardFile(path)
            : LoadCSVFile(path);
        if (records == null || records.Count == 0)
        {
            Console.WriteLine("List is empty.");
            DoExit();
            return;
        }

        //Combine records
        records = CombineRecords(records);

        //Remove duplicates.
        var comparer = new VCardRecordComparer();
        var distinctRecords = records.Distinct(comparer).ToList();

        // Ask if the user wants to save the combined/distinct records as a VCF file
        Console.WriteLine("Do you want to save the records as a VCard file? (y/n)");
        var answerVCF = Console.ReadLine();
        if (answerVCF?.Equals("y", StringComparison.OrdinalIgnoreCase) == true)
        {
            var vcfPath = Path.Combine(Path.GetDirectoryName(path) ?? string.Empty,
                $"{Path.GetFileNameWithoutExtension(path)}___CLEANED.vcf");

            File.WriteAllLines(vcfPath, GetVCFLines(distinctRecords));
            Console.WriteLine($"Saved {distinctRecords.Count} records to {vcfPath}");
        }

        // Ask if the user wants to save the combined/distinct records as a CSV file
        Console.WriteLine("Do you want to save the records as a CSV file? (y/n)");
        var answerCSV = Console.ReadLine();
        if (answerVCF?.Equals("y", StringComparison.OrdinalIgnoreCase) == true)
        {
            var csvPath = Path.Combine(Path.GetDirectoryName(path) ?? string.Empty,
                $"{Path.GetFileNameWithoutExtension(path)}___CLEANED.csv");

            Task.Run(() => SaveCSVFileAsync(distinctRecords, csvPath));
            Console.WriteLine($"CSV file saved to {csvPath}");
        }

        //Bye now.
        DoExit();
    }

    private static void DoExit()
    {
        Console.WriteLine("Press <ENTER> to exit...:)");
        Console.ReadLine();
    }

    private static List<VCardRecord> CombineRecords(List<VCardRecord> records)
    {
        var result = records.Where(r => !string.IsNullOrEmpty(r.FullName) && !string.IsNullOrEmpty(r.Tel))
            .GroupBy(r => r.FullName)
            .Select(g =>
            {
                // Take the first record
                var main = g.First();

                // Find a second record with a different Tel (if any)
                var second = g.Skip(1).FirstOrDefault(r =>
                    !string.IsNullOrEmpty(r.Tel) && r.Tel != main.Tel);

                // If found, store its Tel in Tel2
                if (second != null)
                {
                    if (string.IsNullOrWhiteSpace(main.Tel2))
                    {
                        main.Tel2 = second.Tel;
                    }
                    else if (string.IsNullOrWhiteSpace(main.Tel3))
                    {
                        main.Tel3 = second.Tel;
                    }
                    else
                    {
                        Console.WriteLine($"Due to max numbers, skipped {second.Tel} for {main.FullName}.");
                    }

                    // Find a third record with a different Tel (if any)
                    var third = g.Skip(2).FirstOrDefault(r =>
                        !string.IsNullOrEmpty(r.Tel) && r.Tel != main.Tel);
                    if (third != null && string.IsNullOrWhiteSpace(main.Tel3))
                    {
                        main.Tel3 = third?.Tel;
                    }
                    else if (third != null && !string.IsNullOrWhiteSpace(main.Tel3))
                    {
                        Console.WriteLine($"Due to max numbers, skipped {third.Tel} for {main.FullName}.");
                    }

                    //If tel1 is not a mobile number and tel2 or tel3 are a mobile number, swap them
                    var tel1 = main.Tel ?? string.Empty;
                    var tel2 = main.Tel2 ?? string.Empty;
                    var tel3 = main.Tel3 ?? string.Empty;
                    if (!tel1.StartsWith("05") && tel2.StartsWith("05"))
                    {
                        main.Tel = tel2;
                        main.Tel2 = tel1;
                    }
                    else if (!tel1.StartsWith("05") && tel3.StartsWith("05"))
                    {
                        main.Tel = tel3;
                        main.Tel3 = tel1;
                    }
                }

                return main;
            })
            .ToList();

        return result;
    }

    private static List<VCardRecord>? LoadVCardFile(string filePath)
    {
        try
        {
            var lines = File.ReadAllLines(filePath);
            var records = new List<VCardRecord>();
            VCardRecord? current = null;

            foreach (var line in lines)
            {
                if (line.StartsWith("BEGIN:VCARD"))
                {
                    current = new VCardRecord();
                }
                else if (line.StartsWith('N') && current != null)
                {
                    var lineValue = LineRegex().Match(line).Groups[1].Value.Trim();

                    if (line.Contains("ENCODING=QUOTED-PRINTABLE"))
                    {
                        lineValue = DecodeQuotedPrintable(lineValue);
                    }
                    current.Name = lineValue;
                }
                else if (line.StartsWith("FN") && current != null)
                {
                    var lineValue = LineRegex().Match(line).Groups[1].Value.Trim();

                    if (line.Contains("ENCODING=QUOTED-PRINTABLE"))
                    {
                        lineValue = DecodeQuotedPrintable(lineValue);
                    }
                    current.FullName = lineValue;
                }
                else if (line.StartsWith("TEL") && current != null)
                {
                    var val = GetFixedPhoneNumber(line.Trim());
                    if (string.IsNullOrEmpty(current.Tel))
                    {
                        current.Tel = val;
                    }
                    else if (string.IsNullOrEmpty(current.Tel2))
                    {
                        current.Tel2 = val;
                    }
                    else if (string.IsNullOrEmpty(current.Tel3))
                    {
                        current.Tel3 = val;
                    }
                }
                else if (line.StartsWith("END:VCARD") && current != null)
                {
                    if (string.IsNullOrWhiteSpace(current.Name))
                    {
                        current.Name = current.Tel;
                    }
                    if (string.IsNullOrWhiteSpace(current.FullName))
                    {
                        current.FullName = current.Name ?? current.Tel;
                    }
                    records.Add(current);
                    current = null;
                }
            }

            return records;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Could not load {filePath}: {ex.Message}");
            return null;
        }
    }

    private static List<VCardRecord>? LoadCSVFile(string filePath)
    {
        try
        {
            using CsvDataReader csv = CsvDataReader.Create(filePath);
            var records = csv.GetRecords<VCardRecord>().ToList();
            records.ForEach(r =>
            {
                if (string.IsNullOrWhiteSpace(r.Name))
                {
                    r.Name = r.Tel;
                }
                if (string.IsNullOrWhiteSpace(r.FullName))
                {
                    r.Tel = r.Name ?? r.Tel;
                }
                if (r.Name != null && r.Name.Contains("ENCODING=QUOTED-PRINTABLE"))
                {
                    r.Name = DecodeQuotedPrintable(r.Name);
                }
                if (r.FullName != null && r.FullName.Contains("ENCODING=QUOTED-PRINTABLE"))
                {
                    r.Name = DecodeQuotedPrintable(r.FullName);
                }
                r.Tel = GetFixedPhoneNumber(r.Tel);
                r.Tel2 = GetFixedPhoneNumber(r.Tel2);
                r.Tel3 = GetFixedPhoneNumber(r.Tel3);
            });
            return records;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Could not load CSV file {filePath}: {ex.Message}");
            return null;
        }
    }

    private static string? GetFixedPhoneNumber(string? line)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return null;
        }
        var fixedValue = PhoneRegex().Match(line).Value.Replace("+972", "0").Replace("+", "013");
        if (fixedValue.StartsWith("012"))
        {
            fixedValue = string.Concat("013", fixedValue.AsSpan(3));
        }
        if (fixedValue.StartsWith("00"))
        {
            fixedValue = string.Concat("013", fixedValue.AsSpan(2));
        }
        if (!fixedValue.StartsWith("0"))
        {
            Console.WriteLine($"Invalid phone number: {fixedValue}");
        }
        return fixedValue;
    }

    private static List<string> GetVCFLines(IEnumerable<VCardRecord> records)
    {
        var lines = new List<string>();

        foreach (var record in records)
        {
            lines.Add("BEGIN:VCARD");
            lines.Add("VERSION:3.0");
            if (!string.IsNullOrWhiteSpace(record.Name))
            {
                lines.Add("N:;" + record.Name.Trim());
            }
            if (!string.IsNullOrWhiteSpace(record.FullName))
            {
                lines.Add("FN:" + record.FullName.Trim());
            }
            if (!string.IsNullOrWhiteSpace(record.Tel))
            {
                lines.Add("TEL;TYPE=CELL:" + record.Tel.Trim());
            }
            if (!string.IsNullOrWhiteSpace(record.Tel2))
            {
                lines.Add("TEL;TYPE=HOME:" + record.Tel2.Trim());
            }
            if (!string.IsNullOrWhiteSpace(record.Tel3))
            {
                lines.Add("TEL;TYPE=WORK:" + record.Tel3.Trim());
            }
            lines.Add("END:VCARD");
        }

        return lines;
    }

    private static async Task SaveCSVFileAsync(IEnumerable<VCardRecord> records, string path)
    {
        using var writer = CsvDataWriter.Create(path);
        await writer.WriteAsync(records.AsDataReader());
    }

    private static string DecodeQuotedPrintable(string input)
    {
        var inBytes = Encoding.ASCII.GetBytes(input);
        var outBytes = new byte[inBytes.Length];
        var qpd = new QuotedPrintableDecoder();
        qpd.Decode(inBytes, 0, input.Length, outBytes);
        int i = outBytes.Length - 1;
        while (outBytes[i] == 0)
        {
            --i;
        }
        byte[] outFinalBytes = new byte[i + 1];
        Array.Copy(outBytes, outFinalBytes, i + 1);
        var decoded = Encoding.UTF8.GetString(outFinalBytes);
        return decoded;
    }

    [GeneratedRegex(@"\+?\d+")]
    private static partial Regex PhoneRegex();

    [GeneratedRegex(@".+\:(.+)")]
    private static partial Regex LineRegex();

    public class VCardRecord
    {
        public string? Name { get; set; }
        public string? FullName { get; set; }
        public string? Tel { get; set; }
        public string? Tel2 { get; set; }
        public string? Tel3 { get; set; }
    }

    public class VCardRecordComparer : IEqualityComparer<VCardRecord>
    {
        public bool Equals(VCardRecord? x, VCardRecord? y)
        {
            return x?.Tel == y?.Tel && x?.Tel2 == y?.Tel2 && x?.Tel3 == y?.Tel3;
        }

        public int GetHashCode(VCardRecord obj)
        {
            return obj?.Tel?.GetHashCode() ^ obj?.Tel2?.GetHashCode() ^ obj?.Tel3?.GetHashCode() ?? 0;
        }
    }
}