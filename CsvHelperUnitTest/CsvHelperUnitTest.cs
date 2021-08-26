using System;
using System.Collections.Generic;
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
            };

            using (var writer = new StreamWriter($@"C:\workspace\Company\UK\Troubleshooting\TestSanitizeForInjection_False-{DateTime.Now:yyyyMMdd-HHmmss}.csv"))
            using (var csv = new CsvWriter(writer, config))
            {
                csv.WriteRecords(records);
            }
        }
    }
}
