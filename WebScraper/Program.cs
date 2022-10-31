using HtmlAgilityPack;
using Newtonsoft.Json;
using System.Collections.Concurrent;
using WebScraper;

public class Program
{

    public static async Task Main(string[] args)
    {
        int pageNumber = 1;
        string conbisteamerurl = $"/en/categorie/3621/combi-steamers?page={pageNumber}";
        string response = await Helper.Get(Constants.baseUrl + conbisteamerurl);
        List<ProductModel> parsedData = new List<ProductModel>();
        if (!string.IsNullOrWhiteSpace(response))
        {
            while (!string.IsNullOrWhiteSpace(response))
            {
                var products = await ParseHtml(response);
                if(!products.Any())
                {
                    break;
                }
                parsedData.AddRange(products);
                conbisteamerurl = $"/en/categorie/3621/combi-steamers?page={++pageNumber}";
                response = await Helper.Get(Constants.baseUrl + conbisteamerurl);
            }
            Console.WriteLine("Data Extraction Completed");
            if (parsedData.Any())
            {
                Console.WriteLine("Exporting Data in to Excel");
                Helper.WriteListToExcel(parsedData, Constants.path);    
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
        Console.ReadLine();
    }

    /// <summary>
    /// Parse the Html to return Product Data
    /// </summary>
    /// <param name="html"></param>
    /// <returns></returns>
    private static async Task<List<ProductModel>> ParseHtml(string html)
    {
        HtmlDocument htmlDoc = new();
        htmlDoc.LoadHtml(html);
        ConcurrentBag<ProductModel> products = new ConcurrentBag<ProductModel>();
        List<HtmlNode>? nodes = htmlDoc.DocumentNode.SelectNodes("//a[@class='my-2 fw-bolder text-dark']")?.ToList();

        if (nodes != null)
        {
            Console.WriteLine("Data Extraction Started");

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

    /// <summary>
    /// GetProductData From HtmlNode
    /// </summary>
    /// <param name="node"></param>
    /// <returns></returns>
    private static async Task<ProductModel> GetProductFromNode(HtmlNode node)
    {
        ProductModel product = new()
        {
            ProductUrl = node.Attributes["href"].Value,
            ProductName = node.InnerHtml
        };

        //API Call to fetch the element level information
        string response = await Helper.Get(Constants.baseUrl + product.ProductUrl);

        if (!string.IsNullOrEmpty(response))
        {
            HtmlDocument htmlSubDoc = new();
            htmlSubDoc.LoadHtml(response);

            string? availabilityUrl = htmlSubDoc.DocumentNode.SelectSingleNode("//a[@id='checkInStock']")?.Attributes["data-url"].Value;

            if (availabilityUrl != null)
            {
                GetJsonResponse? availabilityresponse = JsonConvert.DeserializeObject<GetJsonResponse>(await Helper.Get(Constants.baseUrl + availabilityUrl));

                if (availabilityresponse != null && availabilityresponse.Success && availabilityresponse.Data != null)
                {
                    product.Availability = Convert.ToInt32(availabilityresponse?.Data?["inStock"]) == 1 ? "In Stock" : "Out of Stock";
                    product.ProductNumber = availabilityresponse?.Data?["products_id"];
                }
                else
                {
                    product.Availability = Constants.errorMessage;
                }
            }
            else
            {
                product.Availability = Constants.errorMessage;
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

                PostJsonResponse? shippingCostsresponse = JsonConvert.DeserializeObject<PostJsonResponse>(await Helper.Post(Constants.baseUrl + shippingCostsUrl, data));

                product.Shippingcosts = shippingCostsresponse != null && shippingCostsresponse.Success
                    ? Helper.RemoveSpecialCharacters(shippingCostsresponse.Costs)
                    : Constants.errorMessage;
            }
            else
            {
                product.Shippingcosts = Constants.errorMessage;
            }

            product.Deliverytime = htmlSubDoc.DocumentNode.SelectSingleNode("//div[@id='deliveryTime']")?.ChildNodes["a"].ChildNodes[0].InnerText;
            product.Price = htmlSubDoc.DocumentNode.SelectSingleNode("//span[@class='price-formatted h4 fw-bolder']")?.Attributes["data-price"]?.Value;

            if (string.IsNullOrWhiteSpace(product.ProductNumber))
            {
                product.ProductNumber = Helper.RemoveSpecialCharacters(htmlSubDoc.DocumentNode.SelectSingleNode("//span[@class='d-none d-md-inline-block pe-2 border-end']")?.ChildNodes["span"]?.InnerText);
            }
        }
        else
        {
            product.Availability = Constants.errorMessage;
            product.Deliverytime = Constants.errorMessage;
            product.Shippingcosts = Constants.errorMessage;
            product.ProductNumber = Constants.errorMessage;
            product.Price = Constants.errorMessage;
        }
        return product;
    }


















}