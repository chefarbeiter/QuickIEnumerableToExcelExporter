using System;
using System.Collections.Generic;
using System.Linq;
using QuickIEnumerableToExcelExporter;
using Xunit;

namespace UnitTests
{
    public class ValuesReaderTests
    {
        [Fact]
        public void HeaderRowShouldBeCreatedByDefault()
        {
            var metadata = MetadataReader.ReadMetadata(typeof(TestExportItem));
            var enumerable = new List<TestExportItem>();
            var configuration = new ExportConfiguration();

            var valuesReader = new ValuesReader<TestExportItem>(metadata, configuration);
            var values = valuesReader.ReadValues(enumerable);

            Assert.Equal(1, values.Count);
            Assert.True(values[0].IsHeaderRow);
        }

        [Fact]
        public void HeaderRowShouldNotBeCreatedByConfiguration()
        {
            var metadata = MetadataReader.ReadMetadata(typeof(TestExportItem));
            var enumerable = new List<TestExportItem>();
            var configuration = new ExportConfiguration { WriteHeaderRow = false };

            var valuesReader = new ValuesReader<TestExportItem>(metadata, configuration);
            var values = valuesReader.ReadValues(enumerable);

            Assert.Equal(0, values.Count);
        }

        [Fact]
        public void AllRowsShouldBeExported()
        {
            var metadata = MetadataReader.ReadMetadata(typeof(TestExportItem));
            var enumerable = new List<TestExportItem> { new TestExportItem(), new TestExportItem() };
            var configuration = new ExportConfiguration();

            var valuesReader = new ValuesReader<TestExportItem>(metadata, configuration);
            var values = valuesReader.ReadValues(enumerable);

            Assert.Equal(3, values.Count);
            Assert.Equal(2, values.Count(v => !v.IsHeaderRow));
        }

        [Fact]
        public void CorrectValuesShouldBeRead()
        {
            var metadata = MetadataReader.ReadMetadata(typeof(TestExportItem));
            var enumerable = new List<TestExportItem> { new TestExportItem { Name = "Jann", Id = 15 } };
            var configuration = new ExportConfiguration { WriteHeaderRow = false };

            var valuesReader = new ValuesReader<TestExportItem>(metadata, configuration);
            var values = valuesReader.ReadValues(enumerable);
            var row = values.SingleOrDefault();

            Assert.NotNull(row);
            Assert.True(row.Values.Any(v => v.Value.ToString().Equals("Jann")));
            Assert.True(row.Values.Any(v => v.Value is int && (int)v.Value == 15));
        }

        [Fact]
        public void SimpleDisplayMemberPathShouldBeUsed()
        {
            var metadata = MetadataReader.ReadMetadata(typeof(TestExportItem));
            var enumerable = new List<TestExportItem> { new TestExportItem { Timestamp = new DateTime(1961, 2, 25) } };
            var configuration = new ExportConfiguration { WriteHeaderRow = false };

            var valuesReader = new ValuesReader<TestExportItem>(metadata, configuration);
            var values = valuesReader.ReadValues(enumerable);
            var row = values.SingleOrDefault();

            Assert.NotNull(row);
            Assert.True(row.Values.Any(v => v.Value is int && (int)v.Value == 1961));
        }

        [Fact]
        public void NullShouldBeNullByDefault()
        {
            var metadata = MetadataReader.ReadMetadata(typeof(TestExportItem));
            var enumerable = new List<TestExportItem> { new TestExportItem() };
            var configuration = new ExportConfiguration { WriteHeaderRow = false };

            var valuesReader = new ValuesReader<TestExportItem>(metadata, configuration);
            var values = valuesReader.ReadValues(enumerable);
            var row = values.SingleOrDefault();

            Assert.NotNull(row);
            Assert.True(row.Values.Any(v => v.Value == null));
            Assert.False(row.Values.Any(v => v.Value != null && v.Value.ToString() == "NULL"));
        }

        [Fact]
        public void NullShouldBeWrittenAsStringByConfiguration()
        {
            var metadata = MetadataReader.ReadMetadata(typeof(TestExportItem));
            var enumerable = new List<TestExportItem> { new TestExportItem() };
            var configuration = new ExportConfiguration { WriteHeaderRow = false, ExportNullAsString = true };

            var valuesReader = new ValuesReader<TestExportItem>(metadata, configuration);
            var values = valuesReader.ReadValues(enumerable);
            var row = values.SingleOrDefault();

            Assert.NotNull(row);
            Assert.False(row.Values.Any(v => v.Value == null));
            Assert.True(row.Values.Any(v => v.Value.ToString() == "NULL"));
        }
    }
}