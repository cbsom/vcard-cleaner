using System.Text;
using System.Text.RegularExpressions;
using MimeKit.Encodings;
using Sylvan.Data;
using Sylvan.Data.Csv;

namespace VCardFileCleaner
{
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
        private static bool _convertToCSV = false;
        private static bool _convertToVCF = false;

        static void Main(string[] args)
        {
            _convertToCSV = args.Contains("-csv") || args.Contains("--csv");
            _convertToVCF = args.Contains("-vcf") || args.Contains("--vcf");

            if (args.Length < 1)
            {
                Console.WriteLine("Usage: VCardFileCleaner.exe <options> <input file path>");
                Console.WriteLine("Run VCardFileCleaner.exe -help for more information.");
                DoExit();
                return;
            }
            if (args.Contains("-help") || args.Contains("-h"))
            {
                Console.WriteLine("Usage: VCardFileCleaner.exe <options> <input file path>");
                Console.WriteLine("Options:");
                Console.WriteLine("\t-csv: Convert the input file to a CSV file.");
                Console.WriteLine("\t-vcf: Convert the input file to a VCF file.");
                DoExit();
                return;
            }
            var path = args[^1];
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                Console.WriteLine("\nInvalid path: " + path);
                DoExit();
                return;
            }

            var records = Path.GetExtension(path).ToLowerInvariant() == ".vcf"
                ? LoadVCardFile(path)
                : LoadCSVFile(path);
            if (records == null || records.Count == 0)
            {
                Console.WriteLine("\nList is empty.");
                DoExit();
                return;
            }

            if (_convertToCSV)
            {
                var csvPath = Path.Combine(Path.GetDirectoryName(path) ?? string.Empty,
                    $"{Path.GetFileNameWithoutExtension(path)}.csv");
                while (File.Exists(csvPath))
                {
                    csvPath = Path.Combine(Path.GetDirectoryName(path) ?? string.Empty,
                                $"{Path.GetFileNameWithoutExtension(path)}_{DateTime.Now:yyyyMMddHHmmss}.csv");
                }
                Task.Run(() => SaveCSVFileAsync(records, csvPath));
                Console.WriteLine($"CSV file saved to {csvPath}");
                Console.WriteLine($"\nPress <ENTER> to open {csvPath} and exit this program.");
                Console.WriteLine($"Press <CNTL+C> to exit this program now.");
                Console.ReadLine();
                System.Diagnostics.Process.Start("explorer", "\"" + csvPath + "\"");
                return;
            }
            if (_convertToVCF)
            {
                var vcfPath = Path.Combine(Path.GetDirectoryName(path) ?? string.Empty,
                    $"{Path.GetFileNameWithoutExtension(path)}.vcf");

                while (File.Exists(vcfPath))
                {
                    vcfPath = Path.Combine(Path.GetDirectoryName(path) ?? string.Empty,
                                $"{Path.GetFileNameWithoutExtension(path)}_{DateTime.Now:yyyyMMddHHmmss}.vcf");
                }
                File.WriteAllLines(vcfPath, GetVCFLines(records));
                Console.WriteLine($"Saved {records.Count} records to {vcfPath}");
                Console.WriteLine($"\nPress <ENTER> to open {vcfPath} and exit this program.");
                Console.WriteLine($"Press <CNTL+C> to exit this program now.");
                Console.ReadLine();
                System.Diagnostics.Process.Start("explorer", "\"" + vcfPath + "\"");
                return;
            }
            //Combine records
            records = CombineRecords(records);

            //Find duplicate phone numbers
            var duplicates = records.SelectMany(r => new[] { r.Tel, r.Tel2, r.Tel3 })
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .GroupBy(t => t)
                .Where(g => g.Count() > 1)
                .ToList();

            if (duplicates.Count > 0 && !_convertToCSV && !_convertToVCF)
            {
                foreach (var group in duplicates)
                {
                    Console.WriteLine("Duplicate phone number: " + group.Key + " - Found " + group.Count() + " times.");
                }
            }

            //Remove duplicates.
            var comparer = new VCardRecordComparer();
            var distinctRecords = records.Distinct(comparer).ToList();

            // Ask if the user wants to save the combined/distinct records as a VCF file
            Console.WriteLine("\nDo you want to save the combined/distinct records as a VCard file? (y/n)");
            var answerVCF = Console.ReadLine();
            if (answerVCF?.Equals("y", StringComparison.OrdinalIgnoreCase) == true)
            {
                var vcfPath = Path.Combine(Path.GetDirectoryName(path) ?? string.Empty,
                    $"{Path.GetFileNameWithoutExtension(path)}___CLEANED.vcf");

                File.WriteAllLines(vcfPath, GetVCFLines(distinctRecords));
                Console.WriteLine($"Saved {distinctRecords.Count} records to {vcfPath}");
            }

            // Ask if the user wants to save the combined/distinct records as a CSV file
            Console.WriteLine("\nDo you want to save the combined/distinct records as a CSV file? (y/n)");
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
            Console.WriteLine("\nPress <ENTER> to exit...:)");
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
                            main.Tel3 = third.Tel;
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
                    if (line.StartsWith("BEGIN:VCARD")) { current = new VCardRecord(); }
                    else if (line.StartsWith('N') && current != null) { current.Name = GetLineValue(line); }
                    else if (line.StartsWith("FN") && current != null) { current.FullName = GetLineValue(line); }
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
                    else if (line.StartsWith("EMAIL") && current != null) { current.Email = GetLineValue(line); }
                    else if (line.StartsWith("SOURCE") && current != null) { current.SOURCE = GetLineValue(line); }
                    else if (line.StartsWith("KIND") && current != null) { current.KIND = GetLineValue(line); }
                    else if (line.StartsWith("XML") && current != null) { current.XML = GetLineValue(line); }
                    else if (line.StartsWith("NICKNAME") && current != null) { current.NICKNAME = GetLineValue(line); }
                    else if (line.StartsWith("PHOTO") && current != null) { current.PHOTO = GetLineValue(line); }
                    else if (line.StartsWith("BDAY") && current != null) { current.BDAY = GetLineValue(line); }
                    else if (line.StartsWith("ANNIVERSARY") && current != null) { current.ANNIVERSARY = GetLineValue(line); }
                    else if (line.StartsWith("GENDER") && current != null) { current.GENDER = GetLineValue(line); }
                    else if (line.StartsWith("ADR") && current != null) { current.ADR = GetLineValue(line); }
                    else if (line.StartsWith("IMPP") && current != null) { current.IMPP = GetLineValue(line); }
                    else if (line.StartsWith("LANG") && current != null) { current.LANG = GetLineValue(line); }
                    else if (line.StartsWith("TZ") && current != null) { current.TZ = GetLineValue(line); }
                    else if (line.StartsWith("GEO") && current != null) { current.GEO = GetLineValue(line); }
                    else if (line.StartsWith("TITLE") && current != null) { current.TITLE = GetLineValue(line); }
                    else if (line.StartsWith("ROLE") && current != null) { current.ROLE = GetLineValue(line); }
                    else if (line.StartsWith("LOGO") && current != null) { current.LOGO = GetLineValue(line); }
                    else if (line.StartsWith("ORG") && current != null) { current.ORG = GetLineValue(line); }
                    else if (line.StartsWith("MEMBER") && current != null) { current.MEMBER = GetLineValue(line); }
                    else if (line.StartsWith("RELATED") && current != null) { current.RELATED = GetLineValue(line); }
                    else if (line.StartsWith("CATEGORIES") && current != null) { current.CATEGORIES = GetLineValue(line); }
                    else if (line.StartsWith("NOTE") && current != null) { current.NOTE = GetLineValue(line); }
                    else if (line.StartsWith("PRODID") && current != null) { current.PRODID = GetLineValue(line); }
                    else if (line.StartsWith("REV") && current != null) { current.REV = GetLineValue(line); }
                    else if (line.StartsWith("SOUND") && current != null) { current.SOUND = GetLineValue(line); }
                    else if (line.StartsWith("UID") && current != null) { current.UID = GetLineValue(line); }
                    else if (line.StartsWith("CLIENTPIDMAP") && current != null) { current.CLIENTPIDMAP = GetLineValue(line); }
                    else if (line.StartsWith("URL") && current != null) { current.URL = GetLineValue(line); }
                    else if (line.StartsWith("VERSION") && current != null) { current.VERSION = GetLineValue(line); }
                    else if (line.StartsWith("KEY") && current != null) { current.KEY = GetLineValue(line); }
                    else if (line.StartsWith("FBURL") && current != null) { current.FBURL = GetLineValue(line); }
                    else if (line.StartsWith("CALADRURI") && current != null) { current.CALADRURI = GetLineValue(line); }
                    else if (line.StartsWith("CALURI") && current != null) { current.CALURI = GetLineValue(line); }
                    else if (line.StartsWith("BIRTHPLACE") && current != null) { current.BIRTHPLACE = GetLineValue(line); }
                    else if (line.StartsWith("DEATHPLACE") && current != null) { current.DEATHPLACE = GetLineValue(line); }
                    else if (line.StartsWith("DEATHDATE") && current != null) { current.DEATHDATE = GetLineValue(line); }
                    else if (line.StartsWith("EXPERTISE") && current != null) { current.EXPERTISE = GetLineValue(line); }
                    else if (line.StartsWith("HOBBY") && current != null) { current.HOBBY = GetLineValue(line); }
                    else if (line.StartsWith("INTEREST") && current != null) { current.INTEREST = GetLineValue(line); }
                    else if (line.StartsWith("ORG-DIRECTORY") && current != null) { current.ORG_DIRECTORY = GetLineValue(line); }
                    else if (line.StartsWith("CONTACT-URI") && current != null) { current.CONTACT_URI = GetLineValue(line); }
                    else if (line.StartsWith("CREATED") && current != null) { current.CREATED = GetLineValue(line); }
                    else if (line.StartsWith("LANGUAGE") && current != null) { current.LANGUAGE = GetLineValue(line); }
                    else if (line.StartsWith("SOCIALPROFILE") && current != null) { current.SOCIALPROFILE = GetLineValue(line); }
                    else if (line.StartsWith("JSPROP") && current != null) { current.JSPROP = GetLineValue(line); }
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

        private static string GetLineValue(string line)
        {
            var lineValue = LineRegex().Match(line).Groups[1].Value.Trim();
            if (line.Contains("ENCODING=QUOTED-PRINTABLE"))
            {
                lineValue = DecodeQuotedPrintable(lineValue);
            }
            return lineValue;
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
                    if (!string.IsNullOrWhiteSpace(r.Name) && r.Name.Contains("ENCODING=QUOTED-PRINTABLE")) { r.Name = DecodeQuotedPrintable(r.Name); }
                    if (!string.IsNullOrWhiteSpace(r.FullName) && r.FullName.Contains("ENCODING=QUOTED-PRINTABLE")) { r.Name = DecodeQuotedPrintable(r.FullName); }
                    r.Tel = GetFixedPhoneNumber(r.Tel);
                    r.Tel2 = GetFixedPhoneNumber(r.Tel2);
                    r.Tel3 = GetFixedPhoneNumber(r.Tel3);
                    if (!string.IsNullOrWhiteSpace(r.Email) && r.Email.Contains("ENCODING=QUOTED-PRINTABLE")) { r.Email = DecodeQuotedPrintable(r.Email); }
                    if (!string.IsNullOrWhiteSpace(r.SOURCE) && r.SOURCE.Contains("ENCODING=QUOTED-PRINTABLE")) { r.SOURCE = DecodeQuotedPrintable(r.SOURCE); }
                    if (!string.IsNullOrWhiteSpace(r.KIND) && r.KIND.Contains("ENCODING=QUOTED-PRINTABLE")) { r.KIND = DecodeQuotedPrintable(r.KIND); }
                    if (!string.IsNullOrWhiteSpace(r.XML) && r.XML.Contains("ENCODING=QUOTED-PRINTABLE")) { r.XML = DecodeQuotedPrintable(r.XML); }
                    if (!string.IsNullOrWhiteSpace(r.NICKNAME) && r.NICKNAME.Contains("ENCODING=QUOTED-PRINTABLE")) { r.NICKNAME = DecodeQuotedPrintable(r.NICKNAME); }
                    if (!string.IsNullOrWhiteSpace(r.PHOTO) && r.PHOTO.Contains("ENCODING=QUOTED-PRINTABLE")) { r.PHOTO = DecodeQuotedPrintable(r.PHOTO); }
                    if (!string.IsNullOrWhiteSpace(r.BDAY) && r.BDAY.Contains("ENCODING=QUOTED-PRINTABLE")) { r.BDAY = DecodeQuotedPrintable(r.BDAY); }
                    if (!string.IsNullOrWhiteSpace(r.ANNIVERSARY) && r.ANNIVERSARY.Contains("ENCODING=QUOTED-PRINTABLE")) { r.ANNIVERSARY = DecodeQuotedPrintable(r.ANNIVERSARY); }
                    if (!string.IsNullOrWhiteSpace(r.GENDER) && r.GENDER.Contains("ENCODING=QUOTED-PRINTABLE")) { r.GENDER = DecodeQuotedPrintable(r.GENDER); }
                    if (!string.IsNullOrWhiteSpace(r.ADR) && r.ADR.Contains("ENCODING=QUOTED-PRINTABLE")) { r.ADR = DecodeQuotedPrintable(r.ADR); }
                    if (!string.IsNullOrWhiteSpace(r.IMPP) && r.IMPP.Contains("ENCODING=QUOTED-PRINTABLE")) { r.IMPP = DecodeQuotedPrintable(r.IMPP); }
                    if (!string.IsNullOrWhiteSpace(r.LANG) && r.LANG.Contains("ENCODING=QUOTED-PRINTABLE")) { r.LANG = DecodeQuotedPrintable(r.LANG); }
                    if (!string.IsNullOrWhiteSpace(r.TZ) && r.TZ.Contains("ENCODING=QUOTED-PRINTABLE")) { r.TZ = DecodeQuotedPrintable(r.TZ); }
                    if (!string.IsNullOrWhiteSpace(r.GEO) && r.GEO.Contains("ENCODING=QUOTED-PRINTABLE")) { r.GEO = DecodeQuotedPrintable(r.GEO); }
                    if (!string.IsNullOrWhiteSpace(r.TITLE) && r.TITLE.Contains("ENCODING=QUOTED-PRINTABLE")) { r.TITLE = DecodeQuotedPrintable(r.TITLE); }
                    if (!string.IsNullOrWhiteSpace(r.ROLE) && r.ROLE.Contains("ENCODING=QUOTED-PRINTABLE")) { r.ROLE = DecodeQuotedPrintable(r.ROLE); }
                    if (!string.IsNullOrWhiteSpace(r.LOGO) && r.LOGO.Contains("ENCODING=QUOTED-PRINTABLE")) { r.LOGO = DecodeQuotedPrintable(r.LOGO); }
                    if (!string.IsNullOrWhiteSpace(r.ORG) && r.ORG.Contains("ENCODING=QUOTED-PRINTABLE")) { r.ORG = DecodeQuotedPrintable(r.ORG); }
                    if (!string.IsNullOrWhiteSpace(r.MEMBER) && r.MEMBER.Contains("ENCODING=QUOTED-PRINTABLE")) { r.MEMBER = DecodeQuotedPrintable(r.MEMBER); }
                    if (!string.IsNullOrWhiteSpace(r.RELATED) && r.RELATED.Contains("ENCODING=QUOTED-PRINTABLE")) { r.RELATED = DecodeQuotedPrintable(r.RELATED); }
                    if (!string.IsNullOrWhiteSpace(r.CATEGORIES) && r.CATEGORIES.Contains("ENCODING=QUOTED-PRINTABLE")) { r.CATEGORIES = DecodeQuotedPrintable(r.CATEGORIES); }
                    if (!string.IsNullOrWhiteSpace(r.NOTE) && r.NOTE.Contains("ENCODING=QUOTED-PRINTABLE")) { r.NOTE = DecodeQuotedPrintable(r.NOTE); }
                    if (!string.IsNullOrWhiteSpace(r.PRODID) && r.PRODID.Contains("ENCODING=QUOTED-PRINTABLE")) { r.PRODID = DecodeQuotedPrintable(r.PRODID); }
                    if (!string.IsNullOrWhiteSpace(r.REV) && r.REV.Contains("ENCODING=QUOTED-PRINTABLE")) { r.REV = DecodeQuotedPrintable(r.REV); }
                    if (!string.IsNullOrWhiteSpace(r.SOUND) && r.SOUND.Contains("ENCODING=QUOTED-PRINTABLE")) { r.SOUND = DecodeQuotedPrintable(r.SOUND); }
                    if (!string.IsNullOrWhiteSpace(r.UID) && r.UID.Contains("ENCODING=QUOTED-PRINTABLE")) { r.UID = DecodeQuotedPrintable(r.UID); }
                    if (!string.IsNullOrWhiteSpace(r.CLIENTPIDMAP) && r.CLIENTPIDMAP.Contains("ENCODING=QUOTED-PRINTABLE")) { r.CLIENTPIDMAP = DecodeQuotedPrintable(r.CLIENTPIDMAP); }
                    if (!string.IsNullOrWhiteSpace(r.URL) && r.URL.Contains("ENCODING=QUOTED-PRINTABLE")) { r.URL = DecodeQuotedPrintable(r.URL); }
                    if (!string.IsNullOrWhiteSpace(r.VERSION) && r.VERSION.Contains("ENCODING=QUOTED-PRINTABLE")) { r.VERSION = DecodeQuotedPrintable(r.VERSION); }
                    if (!string.IsNullOrWhiteSpace(r.KEY) && r.KEY.Contains("ENCODING=QUOTED-PRINTABLE")) { r.KEY = DecodeQuotedPrintable(r.KEY); }
                    if (!string.IsNullOrWhiteSpace(r.FBURL) && r.FBURL.Contains("ENCODING=QUOTED-PRINTABLE")) { r.FBURL = DecodeQuotedPrintable(r.FBURL); }
                    if (!string.IsNullOrWhiteSpace(r.CALADRURI) && r.CALADRURI.Contains("ENCODING=QUOTED-PRINTABLE")) { r.CALADRURI = DecodeQuotedPrintable(r.CALADRURI); }
                    if (!string.IsNullOrWhiteSpace(r.CALURI) && r.CALURI.Contains("ENCODING=QUOTED-PRINTABLE")) { r.CALURI = DecodeQuotedPrintable(r.CALURI); }
                    if (!string.IsNullOrWhiteSpace(r.BIRTHPLACE) && r.BIRTHPLACE.Contains("ENCODING=QUOTED-PRINTABLE")) { r.BIRTHPLACE = DecodeQuotedPrintable(r.BIRTHPLACE); }
                    if (!string.IsNullOrWhiteSpace(r.DEATHPLACE) && r.DEATHPLACE.Contains("ENCODING=QUOTED-PRINTABLE")) { r.DEATHPLACE = DecodeQuotedPrintable(r.DEATHPLACE); }
                    if (!string.IsNullOrWhiteSpace(r.DEATHDATE) && r.DEATHDATE.Contains("ENCODING=QUOTED-PRINTABLE")) { r.DEATHDATE = DecodeQuotedPrintable(r.DEATHDATE); }
                    if (!string.IsNullOrWhiteSpace(r.EXPERTISE) && r.EXPERTISE.Contains("ENCODING=QUOTED-PRINTABLE")) { r.EXPERTISE = DecodeQuotedPrintable(r.EXPERTISE); }
                    if (!string.IsNullOrWhiteSpace(r.HOBBY) && r.HOBBY.Contains("ENCODING=QUOTED-PRINTABLE")) { r.HOBBY = DecodeQuotedPrintable(r.HOBBY); }
                    if (!string.IsNullOrWhiteSpace(r.INTEREST) && r.INTEREST.Contains("ENCODING=QUOTED-PRINTABLE")) { r.INTEREST = DecodeQuotedPrintable(r.INTEREST); }
                    if (!string.IsNullOrWhiteSpace(r.ORG_DIRECTORY) && r.ORG_DIRECTORY.Contains("ENCODING=QUOTED-PRINTABLE")) { r.ORG_DIRECTORY = DecodeQuotedPrintable(r.ORG_DIRECTORY); }
                    if (!string.IsNullOrWhiteSpace(r.CONTACT_URI) && r.CONTACT_URI.Contains("ENCODING=QUOTED-PRINTABLE")) { r.CONTACT_URI = DecodeQuotedPrintable(r.CONTACT_URI); }
                    if (!string.IsNullOrWhiteSpace(r.CREATED) && r.CREATED.Contains("ENCODING=QUOTED-PRINTABLE")) { r.CREATED = DecodeQuotedPrintable(r.CREATED); }
                    if (!string.IsNullOrWhiteSpace(r.LANGUAGE) && r.LANGUAGE.Contains("ENCODING=QUOTED-PRINTABLE")) { r.LANGUAGE = DecodeQuotedPrintable(r.LANGUAGE); }
                    if (!string.IsNullOrWhiteSpace(r.SOCIALPROFILE) && r.SOCIALPROFILE.Contains("ENCODING=QUOTED-PRINTABLE")) { r.SOCIALPROFILE = DecodeQuotedPrintable(r.SOCIALPROFILE); }
                    if (!string.IsNullOrWhiteSpace(r.JSPROP) && r.JSPROP.Contains("ENCODING=QUOTED-PRINTABLE")) { r.JSPROP = DecodeQuotedPrintable(r.JSPROP); }

                });
                return records;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Could not load CSV file {filePath}: {ex.Message}");
                return null;
            }
        }

        private static string GetFixedPhoneNumber(string? line)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                return "";
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
            if (fixedValue.StartsWith("1") && fixedValue.Length == 11)
            {
                fixedValue = string.Concat("013", fixedValue);
            }
            if (!fixedValue.StartsWith("0"))
            {
                if (fixedValue.Length == 10)
                {
                    fixedValue = string.Concat("0131", fixedValue);
                }
                else if (!_convertToCSV && !_convertToVCF)
                {
                    Console.WriteLine($"Invalid phone number: {fixedValue}");
                }
            }
            return fixedValue ?? "";
        }

        private static List<string> GetVCFLines(IEnumerable<VCardRecord> records)
        {
            var lines = new List<string>();

            foreach (var record in records)
            {
                lines.Add("BEGIN:VCARD");
                lines.Add("VERSION:3.0");
                if (!string.IsNullOrWhiteSpace(record.Name)) { lines.Add("N:;" + record.Name.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.FullName)) { lines.Add("FN:" + record.FullName.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.Tel)) { lines.Add("TEL;TYPE=CELL:" + record.Tel.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.Tel2)) { lines.Add("TEL;TYPE=HOME:" + record.Tel2.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.Tel3)) { lines.Add("TEL;TYPE=WORK:" + record.Tel3.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.Email)) { lines.Add("EMAIL:" + record.Email.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.SOURCE)) { lines.Add("SOURCE:" + record.SOURCE.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.KIND)) { lines.Add("KIND:" + record.KIND.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.XML)) { lines.Add("XML:" + record.XML.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.NICKNAME)) { lines.Add("NICKNAME:" + record.NICKNAME.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.PHOTO)) { lines.Add("PHOTO:" + record.PHOTO.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.BDAY)) { lines.Add("BDAY:" + record.BDAY.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.ANNIVERSARY)) { lines.Add("ANNIVERSARY:" + record.ANNIVERSARY.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.GENDER)) { lines.Add("GENDER:" + record.GENDER.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.ADR)) { lines.Add("ADR:" + record.ADR.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.IMPP)) { lines.Add("IMPP:" + record.IMPP.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.LANG)) { lines.Add("LANG:" + record.LANG.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.TZ)) { lines.Add("TZ:" + record.TZ.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.GEO)) { lines.Add("GEO:" + record.GEO.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.TITLE)) { lines.Add("TITLE:" + record.TITLE.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.ROLE)) { lines.Add("ROLE:" + record.ROLE.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.LOGO)) { lines.Add("LOGO:" + record.LOGO.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.ORG)) { lines.Add("ORG:" + record.ORG.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.MEMBER)) { lines.Add("MEMBER:" + record.MEMBER.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.RELATED)) { lines.Add("RELATED:" + record.RELATED.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.CATEGORIES)) { lines.Add("CATEGORIES:" + record.CATEGORIES.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.NOTE)) { lines.Add("NOTE:" + record.NOTE.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.PRODID)) { lines.Add("PRODID:" + record.PRODID.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.REV)) { lines.Add("REV:" + record.REV.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.SOUND)) { lines.Add("SOUND:" + record.SOUND.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.UID)) { lines.Add("UID:" + record.UID.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.CLIENTPIDMAP)) { lines.Add("CLIENTPIDMAP:" + record.CLIENTPIDMAP.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.URL)) { lines.Add("URL:" + record.URL.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.VERSION)) { lines.Add("VERSION:" + record.VERSION.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.KEY)) { lines.Add("KEY:" + record.KEY.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.FBURL)) { lines.Add("FBURL:" + record.FBURL.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.CALADRURI)) { lines.Add("CALADRURI:" + record.CALADRURI.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.CALURI)) { lines.Add("CALURI:" + record.CALURI.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.BIRTHPLACE)) { lines.Add("BIRTHPLACE:" + record.BIRTHPLACE.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.DEATHPLACE)) { lines.Add("DEATHPLACE:" + record.DEATHPLACE.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.DEATHDATE)) { lines.Add("DEATHDATE:" + record.DEATHDATE.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.EXPERTISE)) { lines.Add("EXPERTISE:" + record.EXPERTISE.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.HOBBY)) { lines.Add("HOBBY:" + record.HOBBY.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.INTEREST)) { lines.Add("INTEREST:" + record.INTEREST.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.ORG_DIRECTORY)) { lines.Add("ORG-DIRECTORY:" + record.ORG_DIRECTORY.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.CONTACT_URI)) { lines.Add("CONTACT-URI:" + record.CONTACT_URI.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.CREATED)) { lines.Add("CREATED:" + record.CREATED.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.LANGUAGE)) { lines.Add("LANGUAGE:" + record.LANGUAGE.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.SOCIALPROFILE)) { lines.Add("SOCIALPROFILE:" + record.SOCIALPROFILE.Trim()); }
                if (!string.IsNullOrWhiteSpace(record.JSPROP)) { lines.Add("JSPROP:" + record.JSPROP.Trim()); }
                lines.Add("END:VCARD");
            }

            return lines;
        }

        private static async Task SaveCSVFileAsync(IEnumerable<VCardRecord> records, string path)
        {
            using var writer = CsvDataWriter.Create(path, new CsvDataWriterOptions
            {
                WriteHeaders = true,
                QuoteStrings = CsvStringQuoting.AlwaysQuote,
                Style = CsvStyle.Standard
            });
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
    }
}