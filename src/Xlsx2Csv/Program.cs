﻿using System.Diagnostics;
using System.Text;
using CommandLine;
using MiniExcelLibs;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
Parser.Default.ParseArguments<Options>(args).WithParsed(XlsxUtility.Convert);


public partial class XlsxUtility
{
    private static readonly char[] sourceArray = [',', '|', '\t'];

    public static void Convert(Options o)
    {
        long startTime = Stopwatch.GetTimestamp();
        var csvs = new List<string>();
        var files = Directory.GetFiles(o.InputFolder, "*.xlsx", SearchOption.AllDirectories).ToList();
        files.RemoveAll(v => v.Contains('~'));
        foreach (var file in files) {
            var inputFile = new FileInfo(file);
            if (!inputFile.Exists) throw new FileNotFoundException("File not exists");

            if (!sourceArray.Contains(o.Delimiter))
                o.Delimiter = ',';

            var worksheetNames = MiniExcel.GetSheetNames(file);

            Console.WriteLine(file);
            Console.WriteLine("------------------------------");
            
            foreach (var sheetName in worksheetNames) {
                var outPath = $"{o.OutputFolder}\\{sheetName}.csv";
                using FileStream xlsx = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using FileStream csv = new FileStream(outPath, FileMode.OpenOrCreate);
                csv.SaveAs(xlsx.Query(excelType: ExcelType.XLSX), false, sheetName, excelType: ExcelType.CSV);
                csvs.Add(outPath);
                Console.WriteLine(outPath);
            }
            Console.WriteLine("\n\n");
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