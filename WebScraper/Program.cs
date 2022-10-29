using ClosedXML.Excel;
using HtmlAgilityPack;
using Newtonsoft.Json;
using System.Collections.Concurrent;
using System.Reflection;
using System.Text.RegularExpressions;
using WebScraper;

public class Program
{
    private const string baseUrl = "https://www.horeca.com";
    private const string errorMessage = "Data Not available";
    public static async Task Main(string[] args)
    {
        int pageNumber = 1;
        string conbisteamerurl = $"/en/categorie/3621/combi-steamers?page={pageNumber}";
        string response = await CallUrl(baseUrl + conbisteamerurl);
        List<ProductModel> parsedData = new();
        if (!string.IsNullOrWhiteSpace(response))
        {
            while (!string.IsNullOrWhiteSpace(response))
            {
                parsedData.AddRange(await ParseHtml(response));
                conbisteamerurl = $"/en/categorie/3621/combi-steamers?page={++pageNumber}";
                response = await CallUrl(baseUrl + conbisteamerurl);
            }
            if (parsedData.Any())
            {
                WriteListToExcel(parsedData);
            }
            else
            {
                Console.WriteLine("No Data to export");
            }
        }
        else
        {
            Console.WriteLine("Either the webpage is down or the url has changed");
        }
        _ = Console.ReadLine();
    }

    private static async Task<List<ProductModel>> ParseHtml(string html)
    {
        HtmlDocument htmlDoc = new();
        htmlDoc.LoadHtml(html);
        ConcurrentBag<ProductModel> products = new();
        List<HtmlNode>? nodes = htmlDoc.DocumentNode.SelectNodes("//a[@class='my-2 fw-bolder text-dark']")?.ToList();

        if (nodes != null)
        {
            ParallelOptions parallelOptions = new()
            {
                MaxDegreeOfParallelism = 5
            };

            await Parallel.ForEachAsync(nodes, parallelOptions, async (node, cancellationToken) =>
            {
                ProductModel product = await GetProductFromNode(node);
                products.Add(product);
            });
        }
        return products.ToList();
    }

    private static async Task<ProductModel> GetProductFromNode(HtmlNode node)
    {
        ProductModel product = new()
        {
            ProductUrl = node.Attributes["href"].Value,
            ProductName = node.InnerHtml
        };

        //API Call to fetch the element level information
        string response = await CallUrl(baseUrl + product.ProductUrl);

        if (!string.IsNullOrEmpty(response))
        {
            HtmlDocument htmlSubDoc = new();
            htmlSubDoc.LoadHtml(response);

            string? availabilityUrl = htmlSubDoc.DocumentNode.SelectSingleNode("//a[@id='checkInStock']")?.Attributes["data-url"].Value;

            if (availabilityUrl != null)
            {
                GetJsonResponse? availabilityresponse = JsonConvert.DeserializeObject<GetJsonResponse>(await CallUrl(baseUrl + availabilityUrl));

                if (availabilityresponse != null && availabilityresponse.Success && availabilityresponse.Data != null)
                {
                    product.Availability = Convert.ToInt32(availabilityresponse?.Data?["inStock"]) == 1 ? "In Stock" : "Out of Stock";
                    product.ProductNumber = availabilityresponse?.Data?["products_id"];
                }
                else
                {
                    product.Availability = errorMessage;
                }
            }
            else
            {
                product.Availability = errorMessage;
            }

            string? shippingCostsUrl = htmlSubDoc.DocumentNode.SelectSingleNode("//div[@class='hide position-absolute shipping-costs-destination mt-2 p-3 me-3']")?.Attributes["data-url"].Value;

            if (shippingCostsUrl != null)
            {
                Dictionary<string, string> data = new()
                {
                    { "quantity", "1" },
                    { "countryId", "150" },//CountryId for netherlands
                    { "postalCode", "" }
                };

                PostJsonResponse? shippingCostsresponse = JsonConvert.DeserializeObject<PostJsonResponse>(await PostUrl(baseUrl + shippingCostsUrl, data));

                product.Shippingcosts = shippingCostsresponse != null && shippingCostsresponse.Success
                    ? RemoveSpecialCharacters(shippingCostsresponse.Costs)
                    : errorMessage;
            }
            else
            {
                product.Shippingcosts = errorMessage;
            }

            product.Deliverytime = htmlSubDoc.DocumentNode.SelectSingleNode("//div[@id='deliveryTime']")?.ChildNodes["a"].ChildNodes[0].InnerText;
            product.Price = htmlSubDoc.DocumentNode.SelectSingleNode("//span[@class='price-formatted h4 fw-bolder']")?.Attributes["data-price"]?.Value;

            if (string.IsNullOrWhiteSpace(product.ProductNumber))
            {
                product.ProductNumber = RemoveSpecialCharacters(htmlSubDoc.DocumentNode.SelectSingleNode("//span[@class='d-none d-md-inline-block pe-2 border-end']")?.ChildNodes["span"]?.InnerText);
            }
        }
        else
        {
            product.Availability = errorMessage;
            product.Deliverytime = errorMessage;
            product.Shippingcosts = errorMessage;
            product.ProductNumber = errorMessage;
            product.Price = errorMessage;
        }
        return product;
    }

    private static async Task<string> CallUrl(string fullUrl)
    {
        try
        {
            HttpClient client = new();
            string response = await client.GetStringAsync(fullUrl);
            return response;
        }
        catch
        {
            return "";
        }
    }


    private static async Task<string> PostUrl(string fullUrl, Dictionary<string, string> body)
    {
        try
        {
            HttpClient client = new();
            HttpResponseMessage responseMessage = await client.PostAsync(fullUrl, new FormUrlEncodedContent(body)); ;
            return await responseMessage.Content.ReadAsStringAsync();
        }
        catch
        {
            return "";
        }
    }


    public static string RemoveSpecialCharacters(string? str)
    {
        return string.IsNullOrWhiteSpace(str) ? string.Empty : Regex.Replace(str, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
    }

    public static void WriteListToExcel(List<ProductModel> list)
    {
        try
        {
            XLWorkbook workbook = new();
            _ = workbook.AddWorksheet("CombiSteamers");
            IXLWorksheet ws = workbook.Worksheet("CombiSteamers");

            PropertyInfo[] properties = list.First().GetType().GetProperties();
            List<string> headerNames = properties.Select(prop => prop.Name).ToList();

            _ = ws.FirstRow().Style.Font.SetBold();
            _ = ws.FirstRow().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
            _ = ws.FirstRow().Style.Fill.SetBackgroundColor(XLColor.Yellow);

            for (int i = 0; i < headerNames.Count; i++)
            {
                ws.Cell(1, i + 1).Value = headerNames[i];
            }

            _ = ws.Cell(2, 1).InsertData(list);
            _ = ws.Columns().AdjustToContents();

            workbook.SaveAs(@"D:\ExportedExcel.xlsx");
        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
        }
    }









}