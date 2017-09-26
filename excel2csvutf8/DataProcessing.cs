using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace excel2csvutf8
{
    class DataProcessing
    {
        private String saveFilePath = "";
        private int codepage = 0;
        private Encoding charEncoder = null;
        private Encoding utf8Encoder = Encoding.UTF8;
        private Encoding defEncoder = Encoding.Default;

        public void parse(String filePath, String saveFilePath, String selectedRegion)
        {
            // set save file path
            this.saveFilePath = saveFilePath;

            // check region
            // reference: https://msdn.microsoft.com/ko-kr/library/windows/desktop/dd317756(v=vs.85).aspx
            switch (selectedRegion)
            {
                case "auto":
                    codepage = 0; // auto is not set codepage
                    break;
                case "dos":
                    codepage = 850; // ibm850, OEM Multilingual Latin 1; Western European (DOS)
                    break;
                case "kr":
                    codepage = 51949; // euc-kr, EUC Korean
                    break;
                case "kr949":
                    codepage = 51949; // ks_c_5601-1987, ANSI/OEM Korean (Unified Hangul Code)
                    break;
                case "jp":
                    codepage = 20932; // euc-jp, Japanese (JIS 0208-1990 and 0212-1990)
                    break;
                case "cn":
                    codepage = 51936; // euc-cn, EUC Simplified Chinese; Chinese Simplified (EUC)
                    break;
                default:
                    codepage = 0; // default is not set codepage
                    break;
            }
            charEncoder = Encoding.GetEncoding(codepage);
            Console.WriteLine("Codepage: " + codepage);

            // parse data
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                int columnSize = 32;

                // write processed rows
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read())
                        {
                            int fieldCount = reader.FieldCount;
                            String[] lineItems = new String[fieldCount];
                            String dataline = "";
                            bool emptyLine = true;

                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                String myType = null;
                                String myString = null;

                                Type objFieldType = reader.GetFieldType(i);
                                if (objFieldType == null)
                                {
                                    myType = "Null";
                                }
                                else
                                {
                                    myType = ((Type)objFieldType).ToString().Split('.').Last();
                                }

                                switch (myType)
                                {
                                    case "Double":
                                        myString = reader.GetDouble(i).ToString();
                                        break;
                                    case "Int":
                                        myString = reader.GetInt32(i).ToString();
                                        break;
                                    case "Bool":
                                        myString = reader.GetBoolean(i).ToString();
                                        break;
                                    case "DateTime":
                                        myString = reader.GetDateTime(i).ToString("yyyy-MM-ddTHH:mm:ssZ");
                                        break;
                                    case "String":
                                        myString = reader.GetString(i).ToString();
                                        break;
                                    case "Null":
                                        myString = null;
                                        break;
                                    default:
                                        myString = null;
                                        break;
                                }

                                if (myString != null)
                                {
                                    if (codepage > 0)
                                    { // only if exists codepage
                                        byte[] bytes = charEncoder.GetBytes(myString);
                                        myString = Encoding.Default.GetString(bytes);
                                    }

                                    if (emptyLine == true)
                                    {
                                        emptyLine = false;
                                    }
                                }
                                else
                                {
                                    myString = "";
                                }

                                myString = myString.Trim(); // trimming string
                                myString = Regex.Replace(myString, @"\t|\n|\r", "");
                                myString = Regex.Replace(myString, @"\s+", " ");

                                lineItems[i] = myString;
                            }

                            // check empty line
                            if (emptyLine == false)
                            {
                                List<String> dataItems = new List<String>();
                                for (int k = 0; k < columnSize; k++)
                                {
                                    if (k < lineItems.Length)
                                    {
                                        if (lineItems[k] == null)
                                        {
                                            dataItems.Add("");
                                        } else if (lineItems[k].Equals(""))
                                        {
                                            dataItems.Add("");
                                        } else
                                        {
                                            dataItems.Add(lineItems[k]);
                                        }
                                    }
                                    else
                                    {
                                        dataItems.Add("");
                                    }
                                }

                                dataline = String.Join(",", dataItems.ToArray());

                                // finally: default -> utf8
                                byte[] bytes = defEncoder.GetBytes(dataline);
                                dataline = utf8Encoder.GetString(bytes);

                                this.fwriteLine(dataline);
                            }
                        }
                    } while (reader.NextResult());
                }
            }
        }

        // Save to CSV (comma seperated)
        public void fwriteLine(String dataline)
        {
            String path = saveFilePath;
            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine(dataline);
                }
            }

            using (StreamWriter sw = File.AppendText(path))
            {
                sw.WriteLine(dataline);
            }
        }
    }
}
