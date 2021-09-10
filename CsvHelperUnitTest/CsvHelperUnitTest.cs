using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using CsvHelper;
using CsvHelper.Configuration;
using NUnit.Framework;

namespace CsvHelperUnitTest
{
    public class Foo
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }

    [TestFixture]
    public class CsvHelperUnitTest
    {
        [Test]
        public void Test20210826_001()
        {
            var config = new CsvConfiguration(new CultureInfo("en-GB"));
            config.SanitizeForInjection = true;
            Console.WriteLine($"SanitizeForInjection = {config.SanitizeForInjection}");
        }


        //https://joshclose.github.io/CsvHelper/examples/writing/write-class-objects/
        [Test]
        public void Test20210826_002()
        {
            var config = new CsvConfiguration(new CultureInfo("en-GB"));
            config.SanitizeForInjection = true;

            var records = new List<Foo>
            {
                new Foo { Id = 1, Name = "one" },
                new Foo { Id = 2, Name = "=AND(1,2)" }, 
                new Foo { Id = 3, Name = "=2+5+cmd|' /C calc'!A0" },//https://www.cnblogs.com/micr067/p/14158456.html
            };
            
            using (var writer = new StreamWriter($@"C:\workspace\Company\UK\Troubleshooting\TestSanitizeForInjection_True-{DateTime.Now:yyyyMMdd-HHmmss}.csv"))
            using (var csv = new CsvWriter(writer,config ))
            {
                csv.WriteRecords(records);
            }
        }

        [Test]
        public void Test20210826_003()
        {
            var config = new CsvConfiguration(new CultureInfo("en-GB"));
            //config.SanitizeForInjection = false;//default value is false

            var records = new List<Foo>
            {
                new Foo { Id = 1, Name = "one" },
                new Foo { Id = 2, Name = "=AND(1,2)" },
                new Foo { Id = 3, Name = "=2+5+cmd|' /C calc'!A0" },//https://www.cnblogs.com/micr067/p/14158456.html
            };

            using (var writer = new StreamWriter($@"C:\workspace\Company\UK\Troubleshooting\TestSanitizeForInjection_False-{DateTime.Now:yyyyMMdd-HHmmss}.csv"))
            using (var csv = new CsvWriter(writer, config))
            {
                csv.WriteRecords(records);
            }
        }

        [Test]
        public void Test20210827_001()
        {
            var config = new CsvConfiguration(new CultureInfo("en-GB"));
            config.SanitizeForInjection = true;

            var records = new List<Foo>
            {
                new Foo { Id = 1, Name = "one" },
                new Foo { Id = 2, Name = "=AND(1,2)" },
                new Foo { Id = 3, Name = "\"123\"" },
                new Foo { Id = 4, Name = "\"=AND(1,2)\"" },
                new Foo { Id = 5, Name = "\"=AND(1,2)" },
            };

            using (var writer = new StreamWriter($@"C:\workspace\Company\UK\Troubleshooting\TestSanitizeForInjection_False-{DateTime.Now:yyyyMMdd-HHmmss}.csv"))
            using (var csv = new CsvWriter(writer, config))
            {
                csv.WriteRecords(records);
            }
        }

        [Test]
        public void Test20210830_001()
        {
            var bytes = Array.Empty<byte>();
            Console.WriteLine($"bytes.Length = {bytes.Length}");
        }


        [Test]
        public void Test20210830_002()
        {
            //https://github.com/JoshClose/CsvHelper/issues/1399
            var dt = new DataTable();
            dt.Columns.Add("Id", typeof(int));
            dt.Columns.Add("Name", typeof(string));

            var row = dt.NewRow();
            row["Id"] = 1;
            row["Name"] = "one";
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Id"] = 2;
            row["Name"] = "two";
            dt.Rows.Add(row);

            using (var writer = new StringWriter())
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                foreach (DataColumn dc in dt.Columns)
                {
                    csv.WriteField(dc.ColumnName);
                }
                csv.NextRecord();

                foreach (DataRow dr in dt.Rows)
                {
                    foreach (DataColumn dc in dt.Columns)
                    {
                        csv.WriteField(dr[dc]);
                    }
                    csv.NextRecord();
                }

                Console.WriteLine(writer.ToString());
            }
        }

        [Test]
        public void Test20210830_003()
        {
            //https://github.com/JoshClose/CsvHelper/issues/1399
            var dt = new DataTable();
            dt.Columns.Add("Id", typeof(int));
            dt.Columns.Add("Name", typeof(string));

            var row = dt.NewRow();
            row["Id"] = 1;
            row["Name"] = "one";
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Id"] = 2;
            row["Name"] = "two";
            dt.Rows.Add(row);

            row = dt.NewRow();
            row["Id"] = 3;
            row["Name"] = "=AND(1,2)";
            dt.Rows.Add(row);

            var config = new CsvConfiguration(CultureInfo.InvariantCulture);
            config.SanitizeForInjection = true;

            using (var writer = new StreamWriter($@"C:\workspace\Company\UK\Troubleshooting\TestSanitizeForInjection_True-{DateTime.Now:yyyyMMdd-HHmmss}.csv"))
            using (var csv = new CsvWriter(writer, config))
            {
                foreach (DataColumn dc in dt.Columns)
                {
                    csv.WriteField(dc.ColumnName);
                }
                csv.NextRecord();

                foreach (DataRow dr in dt.Rows)
                {
                    foreach (DataColumn dc in dt.Columns)
                    {
                        csv.WriteField(dr[dc]);
                    }
                    csv.NextRecord();
                }
                
                csv.Flush();
            }
        }

        [Test]
        public void Test20210831_001()
        {
            //https://github.com/JoshClose/CsvHelper/issues/1570
            var folder = @"C:\workspace\Company\Test\2021\0831";
            var path = Path.Combine(folder, "test.csv");
            var path2 = Path.Combine(folder, "test2.csv");
            var array = File.ReadAllLines(path);

            CsvConfiguration csvConfiguration = new CsvConfiguration(CultureInfo.InvariantCulture);
            csvConfiguration.SanitizeForInjection = true;
            csvConfiguration.HasHeaderRecord = false;

            using (var writer = new StreamWriter(path2, append: false))
            {
                CsvWriter csvWriter = new CsvWriter(writer, csvConfiguration);
                foreach (var field in array)
                {
                    csvWriter.WriteField(field);
                    csvWriter.NextRecord();
                }
                csvWriter.Flush();
                writer.Flush();
            }
        }

        [Test]
        public void Test20210831_002()
        {
            var folder = @"C:\workspace\Company\Test\2021\0831";
            var path = Path.Combine(folder, "test.csv");
            var path2 = Path.Combine(folder, "test2.csv");

            byte[] bytes;
            using (FileStream fs = new FileStream(path, FileMode.Open))
            {
                bytes = new byte[fs.Length];
                fs.Read(bytes, 0, bytes.Length);
            }

            MemoryStream stream = new MemoryStream(bytes);

            CsvConfiguration csvConfiguration = new CsvConfiguration(CultureInfo.InvariantCulture);
            csvConfiguration.SanitizeForInjection = true;
            csvConfiguration.HasHeaderRecord = false;
            csvConfiguration.Mode = CsvMode.NoEscape;

            using (var reader = new StreamReader(stream))
            {
                using (var csvReader = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    bool flag1 = csvReader.Read();
                    bool flag2 = csvReader.ReadHeader();
                    Console.WriteLine($"flag1 = {flag1}, flag2 = {flag2}");

                    Console.WriteLine($"csvReader.HeaderRecord.Length = {csvReader.HeaderRecord.Length}");
                    Console.WriteLine($"csvReader.ColumnCount = {csvReader.ColumnCount}");

                    var count = csvReader.HeaderRecord.Length;
                    for (int i = 0; i < count; i++)
                    {
                        var str = csvReader.GetField(0);
                        Console.WriteLine(str);
                    }
                    Console.WriteLine("Start to read second line");
                    while (csvReader.Read())
                    {
                        for (int i = 0; i < count; i++)
                        {
                            var str = csvReader.GetField(0);
                            Console.WriteLine(str);
                        }
                    }
                }
            }
        }

        [Test]
        public void Test20210909_001()
        {
            char tab = '\t';//(0x09)
            char carriageReturn = '\r';//(0x0D)
            PrintHex(tab);
            PrintHex(carriageReturn);
        }

        private void PrintHex(char character)
        {
            //https://owasp.org/www-community/attacks/CSV_Injection
            byte b = Convert.ToByte(character);
            var hex = b.ToString("X2");
            Console.WriteLine(hex);
        }

        [Test]
        public void Test20210910_001()
        {
            //https://joshclose.github.io/CsvHelper/examples/csvdatareader/
            CsvConfiguration csvConfiguration = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                SanitizeForInjection = true,
                InjectionEscapeCharacter = '\'',
                InjectionCharacters = new[] { '=', '@', '+', '-', '\t', '\r' },
                BadDataFound = null
            };

            var folder = @"C:\workspace\Company\Test\2021\0831";
            var path = Path.Combine(folder, "test.csv");
            DataTable dataTable = new DataTable();
            using (var reader = new StreamReader(path))
            using (var csv = new CsvReader(reader, csvConfiguration))
            {
                // Do any configuration to `CsvReader` before creating CsvDataReader.
                using (var dr = new CsvDataReader(csv))
                {
                    dataTable.Load(dr);
                }
            }
            Console.WriteLine(dataTable.Rows.Count);
        }

    }
}
