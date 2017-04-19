using System;
using System.Linq;
using QuickIEnumerableToExcelExporter;
using Xunit;

namespace UnitTests
{
    public class MetadataReaderTests
    {
        [Fact]
        public void ReadMetadataFromType()
        {
            var metadata = MetadataReader.ReadMetadata(typeof(TestExportItem));

            Assert.Equal(4, metadata.Count);
        }

        [Fact]
        public void ReadHeaderFromPropertyAndAttribute()
        {
            var metadata = MetadataReader.ReadMetadata(typeof(TestExportItem));

            Assert.Equal("Name", metadata[0].Header);
            Assert.Equal("Identity", metadata[1].Header);
        }

        [Fact]
        public void PropertiesWithIgnoreFlagShouldBeIgnored()
        {
            var metadata = MetadataReader.ReadMetadata(typeof(TestExportItemWithIgnore));

            Assert.Equal(2, metadata.Count);
            Assert.False(metadata.Any(m => m.Header == "Password"));
        }

        [Fact]
        public void PropertyWithSimpleDisplayMemberPath()
        {
            var metadata = MetadataReader.ReadMetadata(typeof(TestExportItem));
            var timestampProperty = metadata.Single(m => m.Header == "Timestamp");

            Assert.NotNull(timestampProperty.Property.InnerProperty);
            Assert.Equal("Timestamp", timestampProperty.Property.PropertyInfo.Name);
            Assert.Equal("Year", timestampProperty.Property.InnerProperty.PropertyInfo.Name);
        }

        [Fact]
        public void PropertyWithCompleyDisplayMemberPath()
        {
            var metadata = MetadataReader.ReadMetadata(typeof(TestExportItem));
            var nestedProperty = metadata.Single(m => m.Header == "Nested");

            // First level
            var firstLevel = nestedProperty.Property;
            Assert.NotNull(firstLevel.InnerProperty);
            Assert.Equal("Nested", firstLevel.PropertyInfo.Name);

            // Second level
            var secondLevel = firstLevel.InnerProperty;
            Assert.NotNull(secondLevel.InnerProperty);
            Assert.Equal("Inner", secondLevel.PropertyInfo.Name);

            // Third level
            var thirdLevel = secondLevel.InnerProperty;
            Assert.Null(thirdLevel.InnerProperty);
            Assert.Equal("Value", thirdLevel.PropertyInfo.Name);
        }

        private class TestExportItem
        {
            public string Name { get; set; }

            [ExportToExcel(Header = "Identity")]
            public int Id { get; set; }

            [ExportToExcel(DisplayMemberPath = "Year")]
            public DateTime Timestamp { get; set; }

            [ExportToExcel(DisplayMemberPath = "Inner.Value")]
            public NestedItem Nested { get; set; }
        }

        private class NestedItem
        {
            public string Value { get; set; }

            public NestedItem Inner { get; set; }
        }

        private class TestExportItemWithIgnore
        {
            public int Id { get; set; }

            public string Name { get; set; }

            [ExportToExcel(Ignore = true)]
            public string Password { get; set; }
        }
    }
}