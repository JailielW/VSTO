using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;
using System.Xml.Linq;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void BtnStockPrice_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet wrksht = Globals.ThisAddIn.Application.ActiveSheet;

            Range wrkshtcolumns = wrksht.Columns[1];

            foreach (Range cell in wrkshtcolumns.Cells)
            {
                if (cell.Value2 != null)
                {
                    string base_url = $"https://api.iextrading.com/1.0/stock/{cell.Value2}/quote";
                    
                    try
                    {
                        // Get the web response.
                        string result = GetWebResponse(base_url);
                        var json_serializer = new JavaScriptSerializer();
                        var routes_list = (IDictionary<string, object>)json_serializer.DeserializeObject(result);

                        Range rang = wrksht.Cells[cell.Row, cell.Column + 1];
                        rang.Value2 = routes_list["close"].ToString();
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }

        }

        private string GetWebResponse(string url)
        {
            WebClient web_client = new WebClient();

            Stream response = web_client.OpenRead(url);

            using (StreamReader stream_reader = new StreamReader(response))
            {
                string result = stream_reader.ReadToEnd();
                stream_reader.Close();
                return result;
            }
        }
    }
}
