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

    }
}
