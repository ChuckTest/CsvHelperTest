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
            };
            
            using (var writer = new StreamWriter($@"C:\workspace\Company\UK\Troubleshooting\test-{DateTime.Now:yyyyMMdd-HHmmss}.csv"))
            using (var csv = new CsvWriter(writer,config ))
            {
                csv.WriteRecords(records);
            }
        }
    }
}
