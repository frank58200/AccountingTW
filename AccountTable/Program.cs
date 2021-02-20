using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;
using System.Xml;
using AccountingModel;

namespace AccountTable
{
    class Program
    {
        static void Main(string[] args)
        {

            string FileName = "t_acc_1_o.odt";
            string ContentEntryName = "content.xml";

            var executePath = Assembly.GetEntryAssembly().Location;
            var OutputPath = string.Empty;
#if DEBUG
            OutputPath = Directory.GetParent(executePath).Parent.Parent.Parent.FullName;


#else
            OutputPath = executePath;
#endif

            List<AccountingItem> accountingItems = new List<AccountingItem>();

            using (var file = ZipFile.Open(FileName, ZipArchiveMode.Read))
            {
                var content = file.GetEntry(ContentEntryName);
                if (content != null)
                {
                    XmlDocument xml = new XmlDocument();
                    xml.Load(content.Open());

                    var table = xml.GetElementsByTagName("table:table-row");

                    string Id1st = string.Empty;
                    string Id2nd = string.Empty;
                    string Id3rd = string.Empty;
                    string Id4th = string.Empty;

                    string NameCh = string.Empty;
                    string NameEn = string.Empty;

                    string DescpCh = string.Empty;
                    string DescpEn = string.Empty;
                    string ItemId = string.Empty;


                    foreach (XmlNode row in table)
                    {
                        var account = new AccountingItem();

                        string row_Id1st = row.ChildNodes[0].InnerText;
                        string row_Id2nd = row.ChildNodes[1].InnerText;
                        string row_Id3rd = row.ChildNodes[2].InnerText;
                        string row_Id4th = row.ChildNodes[3].InnerText;



                        string row_NameCh = row.ChildNodes[4].InnerText;
                        string row_NameEn = row.ChildNodes[5].InnerText;

                        string row_DescpCh = row.ChildNodes[6].InnerText;
                        string row_DescpEn = row.ChildNodes[7].InnerText;

                        if (row_Id1st.Contains("一級"))
                        {
                            continue;
                        }

                        if (!string.IsNullOrEmpty(row_Id1st))
                        {
                            Id1st = row_Id1st;
                            Id2nd = string.Empty;
                            Id3rd = string.Empty;
                            Id4th = string.Empty;
                            ItemId = Id1st;

                        }

                        if (!string.IsNullOrEmpty(row_Id2nd))
                        {
                            Id2nd = row_Id2nd;
                            Id3rd = string.Empty;
                            Id4th = string.Empty;
                            ItemId = Id2nd;
                        }
                        if (!string.IsNullOrEmpty(row_Id3rd))
                        {
                            Id3rd = row_Id3rd;
                            Id4th = string.Empty;
                            ItemId = Id3rd;
                        }

                        if (!string.IsNullOrEmpty(row_Id4th))
                        {
                            Id4th = row_Id4th;
                            ItemId = Id4th;
                        }

                        if (!string.IsNullOrEmpty(row_NameCh))
                        {
                            NameCh = row_NameCh;
                            NameEn = row_NameEn;

                        }

                        if (!string.IsNullOrEmpty(row_DescpCh))
                        {
                            DescpCh = row_DescpCh;
                            DescpEn = row_DescpEn;
                        }

                        //Console.WriteLine("{0,5} | {1,5} | {2,5} | {3,5} | {8,5} | {4} | {5} | {6} | {7}", Id1st, Id2nd, Id3rd, Id4th, NameCh, NameEn, DescpCh, DescpEn,ItemId);

                        account.ItemId = ItemId;
                        account.Item1stId = Id1st;
                        account.Item2ndId = Id2nd;
                        account.Item3rdId = Id3rd;
                        account.Item4thId = Id4th;
                        account.AccountNameCh = NameCh;
                        account.AccountNameEn = NameEn;
                        account.DescriptionCh = DescpCh;
                        account.DescriptionEn = DescpEn;

                        accountingItems.Add(account);
                    }
                }
            }

            JsonSerializerOptions options = new JsonSerializerOptions();
            options.Encoder = JavaScriptEncoder.Create(UnicodeRanges.All);
            options.WriteIndented = true;

            var json = JsonSerializer.Serialize(accountingItems, options);


            string AccountFileName = "Accounting.json";
            File.WriteAllText(OutputPath + Path.DirectorySeparatorChar + AccountFileName, json);
            File.WriteAllText(Directory.GetParent(executePath).FullName + Path.DirectorySeparatorChar + AccountFileName, json);
            


        }
    }
}
