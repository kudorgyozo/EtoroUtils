using ClosedXML.Excel;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EtoroUtils {
    public class EtoroStatementProcessor {
        private Dictionary<string, string> cache = new(200);

        string GetCountryNameFromCode(string code) {
            if (cache.ContainsKey(code)) return cache[code];

            try {
                var cultureInfo = new CultureInfo(code);
                var ri = new RegionInfo(cultureInfo.Name);
                cache[code] = ri.EnglishName;

                return ri.EnglishName;
            } catch (ArgumentException) {
                cache[code] = code;

                return code;
            }
        }

        public void Process(string path) {
            using var wb = new XLWorkbook(path);
            GenerateProfitPerTypePerCountry(wb);
            GenerateProfitPerCountry(wb);
            GenerateDividendPerCountry(wb);
            wb.Save();

        }

        void GenerateProfitPerTypePerCountry(XLWorkbook wb) {
            var ws = wb.Worksheet("Closed Positions");

            var firstRow = ws.Row(1);
            var typeCol = firstRow.Cells().First(c => c.GetText() == "Type").Address.ColumnLetter;
            var nameCol = firstRow.Cells().First(c => c.GetText() == "ISIN").Address.ColumnLetter;
            var profitCol = firstRow.Cells().First(c => c.GetText() == "Profit(USD)").Address.ColumnLetter;

            var groupCountryTypeProfit = ws.RowsUsed().Skip(1).Select(row => {
                var type = row.Cell(typeCol).GetValue<string>();
                var isin = row.Cell(nameCol).GetValue<string>();
                var profit = row.Cell(profitCol).GetValue<decimal>();
                var country = string.IsNullOrEmpty(isin) ? "_NA" : isin.Substring(0, 2);
                return new {
                    type,
                    country,
                    profit
                };
            })
            .GroupBy(x => x.country, (country, rows) => {
                return new {
                    country,
                    countries = rows.GroupBy(x => x.type, (type, rows) => new {
                        type,
                        profit = rows.Sum(r => r.profit)
                    }).ToList(),
                };
            }).ToList();
            
            if (ws.Workbook.TryGetWorksheet("Country-Type-Profit", out var wsOut)) wsOut.Delete();
            wsOut = ws.Workbook.AddWorksheet("Country-Type-Profit");

            //country | type | profit
            wsOut.Cell(1, 1).Value = "CountryCode";
            wsOut.Cell(1, 2).Value = "Country";
            wsOut.Cell(1, 3).Value = "Type";
            wsOut.Cell(1, 4).Value = "Profit";
            wsOut.Range("a1:d1").SetAutoFilter(true);

            var row = 2;
            foreach (var grp in groupCountryTypeProfit) {
                foreach (var item in grp.countries) {
                    wsOut.Cell(row, 1).Value = grp.country;
                    wsOut.Cell(row, 2).Value = GetCountryNameFromCode(grp.country);
                    wsOut.Cell(row, 3).Value = item.type;
                    wsOut.Cell(row, 4).Value = item.profit;
                    row++;
                }
            }
        }

        void GenerateProfitPerCountry(XLWorkbook wb) {
            var ws = wb.Worksheet("Closed Positions");

            var firstRow = ws.Row(1);
            var typeCol = firstRow.Cells().First(c => c.GetText() == "Type").Address.ColumnLetter;
            var nameCol = firstRow.Cells().First(c => c.GetText() == "ISIN").Address.ColumnLetter;
            var profitCol = firstRow.Cells().First(c => c.GetText() == "Profit(USD)").Address.ColumnLetter;

            var groupCountryProfit = ws.RowsUsed().Skip(1).Select(row => {
                var type = row.Cell(typeCol).GetValue<string>();
                var isin = row.Cell(nameCol).GetValue<string>();
                var profit = row.Cell(profitCol).GetValue<decimal>();
                var country = string.IsNullOrEmpty(isin) ? "_NA" : isin.Substring(0, 2);
                return new {
                    country,
                    profit
                };
            })
            .GroupBy(x => x.country, (country, rows) => {
                return new {
                    country,
                    profit = rows.Sum(r => r.profit)
                };
            }).ToList();

            if (ws.Workbook.TryGetWorksheet("Country-Profit", out var wsOut)) wsOut.Delete();
            wsOut = ws.Workbook.AddWorksheet("Country-Profit");
            //country | profit

            wsOut.Cell(1, 1).Value = "CountryCode";
            wsOut.Cell(1, 2).Value = "Country";
            wsOut.Cell(1, 3).Value = "Profit";
            wsOut.Range("a1:c1").SetAutoFilter(true);

            var row = 2;
            foreach (var item in groupCountryProfit) {
                wsOut.Cell(row, 1).Value = item.country;
                wsOut.Cell(row, 2).Value = GetCountryNameFromCode(item.country);
                wsOut.Cell(row, 3).Value = item.profit;
                row++;
            }
        }

        void GenerateDividendPerCountry(XLWorkbook wb) {
            var ws = wb.Worksheet("Dividends");

            var firstRow = ws.Row(1);
            var typeCol = firstRow.Cells().First(c => c.GetText() == "Type").Address.ColumnLetter;
            var nameCol = firstRow.Cells().First(c => c.GetText() == "ISIN").Address.ColumnLetter;
            var dividendCol = firstRow.Cells().First(c => c.GetText() == "Net Dividend Received (USD)").Address.ColumnLetter;

            var grpCountryDiv = ws.RowsUsed().Skip(1).Select(row => {
                var type = row.Cell(typeCol).GetValue<string>();
                var isin = row.Cell(nameCol).GetValue<string>();
                var dividend = row.Cell(dividendCol).GetValue<decimal>();
                var country = string.IsNullOrEmpty(isin) ? "_NA" : isin.Substring(0, 2);
                return new {
                    country,
                    dividend
                };
            }).GroupBy(x => x.country, (country, rows) => {
                return new {
                    country,
                    dividend = rows.Sum(r => r.dividend)
                };
            }).ToList();

            if (ws.Workbook.TryGetWorksheet("Country-Dividend", out var wsDiv)) wsDiv.Delete();
            wsDiv = ws.Workbook.AddWorksheet("Country-Dividend");

            wsDiv.Cell(1, 1).Value = "CountryCode";
            wsDiv.Cell(1, 2).Value = "Country";
            wsDiv.Cell(1, 3).Value = "Dividend";
            wsDiv.Range("a1:c1").SetAutoFilter(true);

            var row = 2;
            foreach (var item in grpCountryDiv) {
                wsDiv.Cell(row, 1).Value = item.country;
                wsDiv.Cell(row, 2).Value = GetCountryNameFromCode(item.country);
                wsDiv.Cell(row, 3).Value = item.dividend;
                row++;
            }
        }
    }
}
