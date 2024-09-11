using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using CommandLine;
using Sylvan.Data.Csv;
using Sylvan.Data.Excel;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
Parser.Default.ParseArguments<Options>(args).WithParsed(XlsxUtility.Convert);


public partial class XlsxUtility
{
    private static readonly char[] sourceArray = [',', '|', '\t'];

    public static void Convert(Options o)
    {
        long startTime = Stopwatch.GetTimestamp();
        var csvs = new List<string>();
        var files = Directory.GetFiles(o.InputFolder, "*.xlsx");
        foreach (var file in files) {
            var inputFile = new FileInfo(file);
            if (!inputFile.Exists) throw new FileNotFoundException("File not exists");

            if (!sourceArray.Contains(o.Delimiter))
                o.Delimiter = ',';

            var edr = ExcelDataReader.Create(file, new ExcelDataReaderOptions
            {
                GetErrorAsNull = true
            });

            Console.WriteLine(file);
            Console.WriteLine("------------------------------");
            foreach (var sheetName in edr.WorksheetNames) {
                if (edr.TryOpenWorksheet(sheetName)) {
                    var outPath = $"{o.OutputFolder}\\{sheetName}.csv";
                    using CsvDataWriter cdw = CsvDataWriter.Create(outPath, new CsvDataWriterOptions
                    {
                        Delimiter = o.Delimiter
                    });
                    cdw.Write(edr);
                    csvs.Add(outPath);
                    Console.WriteLine(outPath);
                }
            }
            Console.WriteLine("******************************");
        }

        var gb2312 = Encoding.GetEncoding("gb2312");
        foreach (var file in csvs) {
            var b = Encoding.Convert(Encoding.UTF8, gb2312, Encoding.UTF8.GetBytes(File.ReadAllText(file)));
            File.WriteAllText(file, gb2312.GetString(b), gb2312);
        }
        TimeSpan elapsedTime = Stopwatch.GetElapsedTime(startTime);
        Console.WriteLine($"Converted file in {elapsedTime}");
    }
}

public class Options
{
    [Value(0, Required = true, MetaName = "xlsxfile", HelpText = "xlsx folder path")]
    public string InputFolder { get; set; } = "";

    [Value(1, Required = false, MetaName = "outfile", HelpText = "output csv folder path")]
    public string OutputFolder { get; set; } = "";
    
    [Option('d', "delimiter", Required = false, HelpText = "CSV file separator.")]
    public char Delimiter { get; set; }
}