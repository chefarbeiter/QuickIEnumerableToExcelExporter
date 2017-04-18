using System;

namespace QuickIEnumerableToExcelExporter
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportToExcelAttribute : Attribute
    {
        /*
         * Header, Ignore, Format, DisplayMemberPath
         */
    }
}