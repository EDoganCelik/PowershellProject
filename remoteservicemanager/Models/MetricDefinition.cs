using System;
using System.Collections.Generic;

namespace RemoteServiceManager.Models
{
    public enum MetricScope
    {
        Service = 0,
        Patch = 1,
        Both = 2
    }

    public enum MetricKind
    {
        BuiltIn = 0,
        Registry = 1,
        FileVersion = 2,
        Service = 3,
        WmiProperty = 4
    }

    [Serializable]
    public class MetricDefinition
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string ColumnHeader { get; set; }
        public MetricScope Scope { get; set; }
        public MetricKind Kind { get; set; }
        public bool Enabled { get; set; } = true;
        /// <summary>Registry path / file path / service name / WQL</summary>
        public string Param1 { get; set; } = string.Empty;
        /// <summary>Registry value / WMI property name</summary>
        public string Param2 { get; set; } = string.Empty;

        public bool IsCustom => Kind != MetricKind.BuiltIn;

        public string ResultKey => IsCustom ? ("Custom_" + Id) : Id;

        public override string ToString()
        {
            return IsCustom ? ("[Özel] " + DisplayName) : DisplayName;
        }
    }

    /// <summary>Kullanıcı tercihleri: hangi metrikler açık + özel metrikler.</summary>
    [Serializable]
    public class MetricPreferences
    {
        public List<string> EnabledBuiltInIds { get; set; } = new List<string>();
        public List<MetricDefinition> CustomMetrics { get; set; } = new List<MetricDefinition>();
    }
}
