using System;
using System.Collections.Generic;

namespace RemoteServiceManager.Models
{
    [Serializable]
    public class PresetRecord
    {
        public string Name { get; set; }
        public string Timestamp { get; set; }
        public List<string> Services { get; set; } = new List<string>();
        public string FilePath { get; set; } = string.Empty;
        public string RegPath { get; set; } = string.Empty;
        public string RegValueName { get; set; } = string.Empty;
        public CheckOptions Options { get; set; } = new CheckOptions();
        public MetricPreferences ServiceMetrics { get; set; }
        public MetricPreferences PatchMetrics { get; set; }
    }

    [Serializable]
    public class AppConfigRoot
    {
        public string MachineBinding { get; set; }
        public string Salt { get; set; }
        public string WrappedKey { get; set; }
        public string CipherText { get; set; }
        public string Iv { get; set; }
    }

    /// <summary>
    /// Şifre korumalı taşınabilir yapılandırma (makine bağından bağımsız).
    /// </summary>
    [Serializable]
    public class PortableExportRoot
    {
        public string Format { get; set; } = "RSMPortable";
        public int Version { get; set; } = 1;
        public string Salt { get; set; }
        public string WrappedKey { get; set; }
        public string CipherText { get; set; }
        public string Iv { get; set; }
    }

    [Serializable]
    public class UiPreferences
    {
        public bool ShowServicePanel { get; set; } = true;
        public bool ShowPatchPanel { get; set; } = true;
    }

    [Serializable]
    public class AppConfigPayload
    {
        public List<PresetRecord> Presets { get; set; } = new List<PresetRecord>();
        public string LastSelectedKey { get; set; }
        public CheckOptions DefaultOptions { get; set; } = new CheckOptions();
        public MetricPreferences ServiceMetrics { get; set; }
        public MetricPreferences PatchMetrics { get; set; }
        public UiPreferences Ui { get; set; } = new UiPreferences();
    }
}
