using System;
using System.IO;
using Microsoft.Extensions.Configuration;
using McMaster.Extensions.CommandLineUtils;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using System.Threading;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Customsearch.v1;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;

namespace pkmncube {
    class Program {
        IConfiguration Configuration { get; set; }
        SpreadsheetsResource GoogleSheets { get; set; }
        CseResource GoogleSearch { get; set; }

        void BuildConfiguration(string[] args) {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("config.json", true)
                .AddCommandLine(args);

            Configuration = builder.Build();
        }

        void Configure(string[] args) {
            var app = new CommandLineApplication();

            var googleId = app.Option("--google-client-id <KEY>", "Required. The Google client ID.", CommandOptionType.SingleValue);
            var googleSecret = app.Option("--google-client-secret <KEY>", "Required. The Google client secret.", CommandOptionType.SingleValue);
            var googleKey = app.Option("--google-custom-search-api-key <KEY>", "Required. The API key for Google Custom Search.", CommandOptionType.SingleValue);
            var googleCx = app.Option("--google-custom-search-cx <CX>", "Required. The search engine for Google Custom Search.", CommandOptionType.SingleValue);
            var sheet = app.Option("--sheet <ID>", "Required. The spreadsheet ID to operate on.", CommandOptionType.SingleValue);

            var help = app.HelpOption();

            var defaultHandler = app.ValidationErrorHandler;
            app.ValidationErrorHandler = arg => { defaultHandler(arg); Environment.Exit(1); return 1; };

            BuildConfiguration(args);

            if (Configuration["google-client-id"] == null) googleId.IsRequired();
            if (Configuration["google-client-secret"] == null) googleSecret.IsRequired();
            if (Configuration["google-custom-search-api-key"] == null) googleKey.IsRequired();
            if (Configuration["google-custom-search-cx"] == null) googleCx.IsRequired();
            if (Configuration["sheet"] == null) sheet.IsRequired();

            try {
                app.Execute(args);
            } catch (CommandParsingException e) {
                var color = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Error.WriteLine(e.Message);
                Console.ForegroundColor = color;
                Environment.Exit(1);
            }

            if (help.HasValue()) Environment.Exit(0);
        }

        string TCGPlayerSlug(string set) => set.ToLower().Replace(" ", "-").Replace("'", "").Replace(".", "");

        async Task GoogleLogin() {
            var credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    new ClientSecrets {
                        ClientId = Configuration["google-client-id"],
                        ClientSecret = Configuration["google-client-secret"]
                    },
                    new[] { SheetsService.Scope.Drive },
                    "user", CancellationToken.None);

            var sheetsService = new SheetsService(new BaseClientService.Initializer {
                ApplicationName = "pkmncube",
                HttpClientInitializer = credential
            });

            var customsearchService = new CustomsearchService(new BaseClientService.Initializer {
                ApplicationName = "pkmncube",
                ApiKey = Configuration["google-custom-search-api-key"]
            });

            GoogleSheets = new SpreadsheetsResource(sheetsService);
            GoogleSearch = new CseResource(customsearchService);
        }

        async Task Execute(string[] args) {
            Configure(args);
            await GoogleLogin();
            var spreadsheet = await GoogleSheets.GetByDataFilter(new GetSpreadsheetByDataFilterRequest {
                IncludeGridData = true
            }, Configuration["sheet"]).ExecuteAsync();
            var sheet = spreadsheet.Sheets[0];
            var cells = sheet.Data[0];

            int currentRowNum = 0;
            List<Task<Request>> requests = new List<Task<Request>>();

            foreach (var row in cells.RowData) {
                currentRowNum++;
                int rowNum = currentRowNum;
                Func<Task<Request>> func = async () => {
                    if (rowNum < 2) return null;
                    var name = row.Values[0].FormattedValue;
                    if (name == "" || name == null) return null;
                    var number = row.Values[1].FormattedValue;
                    if (number == null) number = "";
                    var set = row.Values[2].FormattedValue;
                    if (set == null) set = "";

                    Console.WriteLine($"card {rowNum - 1}...");

                    var oldLink = row.Values[5].FormattedValue;
                    if (oldLink != "" && oldLink != null) return null;

                    var query = $"buy \"{name}\" set {set} {number} pokemon";

                    var request = GoogleSearch.List(query);
                    request.Cx = Configuration["google-custom-search-cx"];

                    var list = await request.ExecuteAsync();
                    var attempts = 0;

                    var result = "";

                    if (list?.Items == null) return null;

                    foreach (var possibleResult in list.Items) {
                        var str = possibleResult.Link;
                        attempts++;
                        if (attempts > 5) break;
                        if (!str.Contains(TCGPlayerSlug(set))) continue;
                        if (str.Contains("deck") || str.Contains("product") || str.Contains("price-guide") || str.Contains("secret-rare")) continue;
                        if (set != "" && new Regex(".*-[ -km-uw-~]+[0-9]+$").Match(str).Success && !set.ToLower().Contains("promo")) continue;
                        if (str.Contains(TCGPlayerSlug(name))) { result = str; break; }
                    }

                    foreach (var possibleResult in list.Items) {
                        var str = possibleResult.Link;
                        if (str.Contains("deck") || str.Contains("product") || str.Contains("price-guide") || str.Contains("secret-rare")) continue;
                        if (set != "" && new Regex(".*-[ -km-uw-~]+[0-9]+$").Match(str).Success && !set.ToLower().Contains("promo")) continue;
                        if (str.Contains(TCGPlayerSlug(name))) { result = str; break; }
                    }

                    if (number != "" && !result.EndsWith(number.ToString()) && !(new Regex(".*-lv[0-9]+$")).Match(result).Success) {
                        result = (new Regex("[0-9]+$")).Replace(result, number.ToString());
                    }

                    var link = new CellData();

                    link.UserEnteredValue = new ExtendedValue { StringValue = result };

                    var valueUpdateRequest = new Request {
                        UpdateCells = new UpdateCellsRequest {
                            Fields = "*",
                            Start = new GridCoordinate {
                                ColumnIndex = 5,
                                RowIndex = rowNum - 1
                            },
                            Rows = new[] {new RowData {
                                           Values = new[] {link}
                                       }}
                        }
                    };

                    return valueUpdateRequest;
                };

                requests.Add(func());
            };

            if (requests.Count > 0) {
                var sheetUpdateRequest = new BatchUpdateSpreadsheetRequest {
                    Requests = (await Task.WhenAll<Request>(requests)).Where(e => e != null).ToList()
                };

                await GoogleSheets.BatchUpdate(sheetUpdateRequest, Configuration["sheet"]).ExecuteAsync();
            }
        }

        static void Main(string[] args) {
            (new Program()).Execute(args).Wait();
        }
    }
}
