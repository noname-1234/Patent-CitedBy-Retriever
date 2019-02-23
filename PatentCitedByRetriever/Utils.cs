using ClosedXML.Excel;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace PatentCitedByRetriever
{
    public class Utils
    {
        private const string USER_AGENT = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/36.0.1985.125 Safari/537.36";

        public static string API(string method, string uri, CookieContainer cc = null, string data = null)
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(uri);
            req.Method = method;
            req.UserAgent = USER_AGENT;
            req.KeepAlive = true;

            if (cc != null)
            {
                req.CookieContainer = cc;
            }

            if (method == "POST" && data != null)
            {
                byte[] postData = Encoding.ASCII.GetBytes(data);
                req.ContentType = "application/x-www-form-urlencoded";
                req.ContentLength = data.Length;
                using (Stream stream = req.GetRequestStream())
                {
                    stream.Write(postData, 0, postData.Length);
                }
            }

            HttpWebResponse resp = (HttpWebResponse)req.GetResponse();
            if (resp.StatusCode != HttpStatusCode.OK)
            {
                Debug.WriteLine($"status code not OK: {resp.StatusCode}: {resp.StatusDescription}"); ;
                throw new Exception("API 錯誤");
            }

            string responseString = new StreamReader(resp.GetResponseStream()).ReadToEnd();
            //Debug.WriteLine(responseString);

            resp.Close();
            resp.Dispose();

            return responseString;
        }

        public static List<HtmlNode> findNodesContainKeyWords(HtmlNode node, string keyword)
        {
            List<HtmlNode> ret = new List<HtmlNode>();

            void _findNodesContainKeyWords(HtmlNode n)
            {
                foreach (HtmlNode cn in n.ChildNodes)
                {
                    if (cn.InnerHtml.Contains(keyword))
                    {
                        _findNodesContainKeyWords(cn);
                    }
                }

                if (!n.HasChildNodes && n.InnerHtml.Contains(keyword))
                {
                    ret.Add(n);
                }
            }
            _findNodesContainKeyWords(node);

            return ret;
        }

        public static DataTable HtmlTableNodeToDataTable(HtmlNode node)
        {
            if (node.Name != "table")
            {
                throw new ArgumentException("Only accept table type of HTML node");
            }

            DataTable dt = new DataTable();

            HtmlNode thead = node.Descendants("thead").First();
            HtmlNode thtr = thead.Descendants("tr").First();

            foreach (HtmlNode th in thtr.Descendants("th"))
            {
                dt.Columns.Add(th.InnerText);
            }

            HtmlNode tbody = node.Descendants("tbody").First();

            foreach (HtmlNode tr in tbody.Descendants("tr"))
            {
                DataRow row = dt.NewRow();
                int c = 0;
                foreach (HtmlNode td in tr.Descendants("td"))
                {
                    row[c] = findText(td);
                    c++;
                }
                dt.Rows.Add(row);
            }

            return dt;
        }

        private static string findText(HtmlNode node)
        {
            foreach (HtmlNode cn in node.ChildNodes)
            {
                string innerText = cn.InnerText;
                if (cn.Name == "#text" && string.Empty != Regex.Replace(innerText, @"\s+", string.Empty))
                {
                    return innerText;
                }
                if (cn.HasChildNodes)
                {
                    string childInnerText = findText(cn);
                    if (childInnerText != string.Empty)
                    {
                        return childInnerText;
                    }
                }
            }

            return "";
        }

        public static DataTable retrieveInfo(DataTable dt)
        {
            DataTable ret = new DataTable();
            ret.Columns.Add("Publication number");
            ret.Columns.Add("Publication year");

            foreach (DataRow row in dt.Rows)
            {
                DataRow nr = ret.NewRow();
                nr[0] = row[0];
                if (row[2].ToString() == "")
                {
                    continue;
                }
                nr[1] = DateTime.Parse(row[2].ToString()).Year;
                ret.Rows.Add(nr);
            }

            return ret;
        }

        public static DataTable retrieveDataFromImportedDataTable(DataTable dt, int columnIdx, Action<string> logger = null)
        {
            int minY = int.MaxValue;
            int maxY = int.MinValue;

            if (logger == null)
            {
                logger = (string text) => Debug.WriteLine(text);
            }

            DataColumn col = dt.Columns[columnIdx];
            string author = col.ColumnName;

            logger($"Processing author {author}");

            DataTable dtAuthor = new DataTable
            {
                TableName = author
            };
            dtAuthor.Columns.Add("PN");
            dtAuthor.Columns.Add("No. of Cited By");

            List<Dictionary<int, int>> stList = new List<Dictionary<int, int>>();
            List<string> pnList = new List<string>();
            foreach (DataRow row in dt.Rows)
            {
                if (row[columnIdx].ToString() == "")
                {
                    break;
                }

                string PN = row[columnIdx].ToString();

                logger($"PN: {PN}");

                DataTable ptInfo = getPatentCitedInfo(PN);
                Dictionary<int, int> statistic = statisticPatentInfo(ptInfo);
                stList.Add(statistic);
                pnList.Add(PN);

                if (statistic.Count > 0)
                {
                    int _minY = statistic.Select(x => x.Key).ToList().OrderBy(x => x).First();
                    int _maxY = statistic.Select(x => x.Key).ToList().OrderByDescending(x => x).First();

                    if (minY > _minY)
                    {
                        minY = _minY;
                    }

                    if (maxY < _maxY)
                    {
                        maxY = _maxY;
                    }
                }
            }

            if (stList.Count == 0)
            {
                return dtAuthor;
            }

            foreach (int y in Enumerable.Range(minY, maxY - minY + 1))
            {
                dtAuthor.Columns.Add(y.ToString());
            }

            for (int i = 0; i < stList.Count; i++)
            {
                DataRow row = dtAuthor.NewRow();

                Dictionary<int, int> statistic = stList[i];
                string PN = pnList[i];

                row[0] = PN;
                row[1] = statistic.Select(x => x.Value).Sum();

                for (int cc = 2; cc < dtAuthor.Columns.Count; cc++)
                {
                    int y = int.Parse(dtAuthor.Columns[cc].ColumnName);
                    if (statistic.ContainsKey(y))
                    {
                        row[cc] = statistic[y];
                    }
                    else
                    {
                        row[cc] = 0;
                    }
                }

                dtAuthor.Rows.Add(row);
            }

            return dtAuthor;
        }

        public static Dictionary<int, int> statisticPatentInfo(DataTable dt)
        {
            Dictionary<int, int> ret = new Dictionary<int, int>();

            foreach (DataRow row in dt.Rows)
            {
                int y = int.Parse(row[1].ToString());
                if (ret.ContainsKey(y))
                {
                    ret[y] = ret[y] + 1;
                }
                else
                {
                    ret[y] = 1;
                }
            }

            return ret;
        }

        public static DataTable getPatentCitedInfo(string PN)
        {
            string respText = API("GET", $"https://patents.google.com/patent/US{PN}");
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(respText);

            HtmlNode familySection = doc.DocumentNode
                                     .ChildNodes
                                     .Descendants("section")
                                     .Where(x => x.Attributes["itemprop"] != null && x.Attributes["itemprop"].Value == "family")
                                     .First();

            HtmlNode tblCitedByNode = findNodesContainKeyWords(familySection, "Cited By").FirstOrDefault();
            if (tblCitedByNode != null)
            {
                tblCitedByNode = tblCitedByNode.Ancestors("h2").First();
                while (tblCitedByNode.Name != "table")
                {
                    tblCitedByNode = tblCitedByNode.NextSibling;
                }
            }

            HtmlNode tblFamilyCitingNode = findNodesContainKeyWords(familySection, "Families Citing this family").FirstOrDefault();
            if (tblFamilyCitingNode != null)
            {
                tblFamilyCitingNode = tblFamilyCitingNode.Ancestors("h2").First();
                while (tblFamilyCitingNode.Name != "table")
                {
                    tblFamilyCitingNode = tblFamilyCitingNode.NextSibling;
                }
            }

            DataTable dtCitedBy = new DataTable();
            if (tblCitedByNode != null)
            {
                dtCitedBy = HtmlTableNodeToDataTable(tblCitedByNode);
            }

            DataTable dtFamilyCiting = new DataTable();
            if (tblFamilyCitingNode != null)
            {
                dtFamilyCiting = HtmlTableNodeToDataTable(tblFamilyCitingNode);
            }

            DataTable dtFinal = new DataTable();

            if (dtCitedBy.Rows.Count == 0 && dtFamilyCiting.Rows.Count != 0)
            {
                dtFinal = dtFamilyCiting;
            }
            else if (dtCitedBy.Rows.Count != 0 && dtFamilyCiting.Rows.Count == 0)
            {
                dtFinal = dtCitedBy;
            }
            else if (dtCitedBy.Rows.Count != 0 && dtFamilyCiting.Rows.Count != 0)
            {
                dtCitedBy.Merge(dtFamilyCiting);
                dtFinal = dtCitedBy;
            }

            Debug.WriteLine($"{PN} is retrieved");
            return retrieveInfo(dtFinal);
        }

        public static DataTable getDataTableFromExcel(string filePath, int tab)
        {
            using (XLWorkbook workBook = new XLWorkbook(filePath))
            {
                IXLWorksheet workSheet = workBook.Worksheet(tab);
                DataTable dt = new DataTable();

                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        dt.Rows.Add();
                        int i = 0;

                        IXLCells a = row.Cells();

                        foreach (IXLCell cell in row.Cells(1, row.LastCellUsed().Address.ColumnNumber))
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            i++;
                        }
                    }
                }
                return dt;
            }
        }
    }
}