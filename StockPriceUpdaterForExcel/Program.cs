using System;
using System.Threading.Tasks;
using YahooFinanceApi;
using ClosedXML.Excel;

string filePath = Util.GetArgumentValue(args, "-file");
string sheetName = Util.GetArgumentValue(args, "-sheet");
string tickerTopCell = "C4";
string priceTopCell = "G4";
string sharesOutstandingTopCell = "I4";
string epsTopCell = "K4"; 
string dividendPriceTopCell = "L4";

if (filePath == string.Empty)
{
    Console.WriteLine("Please specify the file path with -file option.");
    return;
}
if (sheetName == string.Empty)
{
    Console.WriteLine("Please specify the sheet name with -sheet option.");
    return;
}

if (File.Exists(filePath))
{
}
else
{
    Console.WriteLine($"File not found. : {filePath}");
    return;
}

var workbook = new XLWorkbook(filePath);
var worksheet = workbook.Worksheet(sheetName);

if(worksheet == null)
{
    Console.WriteLine($"Sheet not found. : {sheetName}");
    return;
}

var tickerCell = worksheet.Cell(tickerTopCell);
var priceCell = worksheet.Cell(priceTopCell);
var epsCell = worksheet.Cell(epsTopCell);
var sharesOutstandingCell = worksheet.Cell(sharesOutstandingTopCell);
var dividendYieldCell = worksheet.Cell(dividendPriceTopCell);

while (true)
{
    var ticker = tickerCell.GetString();
    if (string.IsNullOrEmpty(ticker))
    {
        break;
    }
    double price = await StockPriceFetcher.GetStockPriceAsync(ticker);
    if (price == -1)
    {
        Console.WriteLine($"{ticker}の株価の取得に失敗しました。");
    }
    double sharesOutstanding = await StockPriceFetcher.GetSharesOutstandingAsync(ticker);
    if (sharesOutstanding == -1)
    {
        Console.WriteLine($"{ticker}の発行済み株式数の取得に失敗しました。");
    }
    double eps = await StockPriceFetcher.GetEpsAsync(ticker);
    if (eps == -1)
    {
        Console.WriteLine($"{ticker}のEPSの取得に失敗しました。");
    }
    double dividendYield = await StockPriceFetcher.GetDividendYieldAsync(ticker);
    if (dividendYield == -1)
    {
        Console.WriteLine($"{ticker}の予想配当の取得に失敗しました。");
    }

    priceCell.Value = price;
    sharesOutstandingCell.Value = sharesOutstanding;
    epsCell.Value = eps;
    dividendYieldCell.Value = dividendYield;
    Console.WriteLine($"{ticker} : {price},{sharesOutstanding},{eps},{dividendYield}");

    tickerCell = tickerCell.CellBelow();
    priceCell = priceCell.CellBelow();
    sharesOutstandingCell = sharesOutstandingCell.CellBelow();
    epsCell = epsCell.CellBelow();
    dividendYieldCell = dividendYieldCell.CellBelow();

    // 1秒待機（API制限回避のため）
    await Task.Delay(1000);
}

workbook.Save();
public class StockPriceFetcher
{
    /// <summary>
    /// 指定された銘柄コードの現在の株価を取得します。
    /// </summary>
    /// <param name="ticker">銘柄コード（例: "AAPL", "7203.T"）</param>
    /// <returns>株価（double）</returns>
    public static async Task<double> GetStockPriceAsync(string ticker)
    {
        try
        {
            var securities = await Yahoo.Symbols(ticker).Fields(YahooFinanceApi.Field.RegularMarketPrice).QueryAsync();
            var security = securities[ticker];
            return (double)security[YahooFinanceApi.Field.RegularMarketPrice];
        }
        catch (Exception ex)
        {
            Console.WriteLine($"エラー: {ex.Message}");
            return -1;
        }
    }
    public static async Task<double> GetSharesOutstandingAsync(string ticker)
    {
        try
        {
            var securities = await Yahoo.Symbols(ticker).Fields(YahooFinanceApi.Field.SharesOutstanding).QueryAsync();
            var security = securities[ticker];
            return (double)security[YahooFinanceApi.Field.SharesOutstanding];
        }
        catch (Exception ex)
        {
            Console.WriteLine($"エラー: {ex.Message}");
            return -1;
        }
    }
    public static async Task<double> GetEpsAsync(string ticker)
    {
        try
        {
            var securities = await Yahoo.Symbols(ticker).Fields(YahooFinanceApi.Field.EpsTrailingTwelveMonths).QueryAsync();
            var security = securities[ticker];
            return (double)security[YahooFinanceApi.Field.EpsTrailingTwelveMonths];
        }
        catch (Exception ex)
        {
            Console.WriteLine($"エラー: {ex.Message}");
            return -1;
        }
    }
    public static async Task<double> GetDividendYieldAsync(string ticker)
    {
        try
        {
            var securities = await Yahoo.Symbols(ticker).Fields(YahooFinanceApi.Field.TrailingAnnualDividendRate).QueryAsync();
            var security = securities[ticker];
            return (double)security[YahooFinanceApi.Field.TrailingAnnualDividendRate];
        }
        catch (Exception ex)
        {
            Console.WriteLine($"エラー: {ex.Message}");
            return -1;
        }
    }
}

public class Util
{
    public static bool ArgumentExists(string[] args, string flag)
    {
        bool result = args.Where(x => x.StartsWith("-")).Any(x => x == flag);
        return result;
    }
    public static string GetArgumentValue(string[] args, string key)
    {
        for (int i = 0; i < args.Length; ++i)
        {
            if (args[i].StartsWith("-"))
            {

            }
            else
            {
                continue;
            }

            if (args[i] == key)
            {
            }
            else
            {
                continue;
            }
            
            if (i + 1 < args.Length)
            {
                return args[i + 1];
            }
        }
        return string.Empty;
    }
}