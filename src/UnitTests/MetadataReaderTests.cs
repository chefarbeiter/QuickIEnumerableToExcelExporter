﻿using System.Linq;
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

        [Fact]
        public void ReadStringFormatFromAttribute()
        {
            var metadata = MetadataReader.ReadMetadata(typeof(TestExportItem));

            var idProperty = metadata.Single(m => m.Header == "Identity");
            var nameProperty = metadata.Single(m => m.Header == "Name");

            Assert.Equal("#,#0", idProperty.Format);
            Assert.Null(nameProperty.Format);
        }
    }
}