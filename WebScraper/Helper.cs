using ClosedXML.Excel;
using System.Reflection;
using System.Text.RegularExpressions;

namespace WebScraper
{
    public static class Helper
    {

        /// <summary>
        /// Write data to Excel File
        /// </summary>
        /// <param name="list"></param>
        public static void WriteListToExcel(List<ProductModel> list, string path)
        {
            try
            {
                XLWorkbook workbook = new XLWorkbook();
                workbook.AddWorksheet("CombiSteamers");
                IXLWorksheet ws = workbook.Worksheet("CombiSteamers");

                PropertyInfo[] properties = list.First().GetType().GetProperties();
                List<string> headerNames = properties.Select(prop => prop.Name).ToList();

                ws.FirstRow().Style.Font.SetBold();
                ws.FirstRow().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
                ws.FirstRow().Style.Fill.SetBackgroundColor(XLColor.Yellow);

                for (int i = 0; i < headerNames.Count; i++)
                {
                    ws.Cell(1, i + 1).Value = headerNames[i];
                }

                ws.Cell(2, 1).InsertData(list);
                ws.Columns().AdjustToContents();

                workbook.SaveAs(path);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        /// <summary>
        /// Remove special characters from string
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string RemoveSpecialCharacters(string? str)
        {
            return string.IsNullOrWhiteSpace(str) ? string.Empty : Regex.Replace(str, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
        }

        /// <summary>
        /// Post API Call
        /// </summary>
        /// <param name="fullUrl"></param>
        /// <param name="body"></param>
        /// <returns></returns>
        public static async Task<string> Post(string fullUrl, Dictionary<string, string> body)
        {
            try
            {
                HttpClient client = new();
                HttpResponseMessage responseMessage = await client.PostAsync(fullUrl, new FormUrlEncodedContent(body));
                return await responseMessage.Content.ReadAsStringAsync();
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// Get API Call
        /// </summary>
        /// <param name="fullUrl"></param>
        /// <returns></returns>
        public static async Task<string> Get(string fullUrl)
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
    }
}
