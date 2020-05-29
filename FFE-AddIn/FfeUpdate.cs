using Newtonsoft.Json.Linq;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using Forms = System.Windows.Forms;

namespace FFE
{
    public static class FfeUpdate
    {
        public static JArray GitHubReleases
        {
            get
            {
                string json = FfeWeb.GetHttpResponseContent(FfeSetting.Default.UpdateUrl, true);
                return JArray.Parse(json);
            }
        }

        public static (bool isNewVersionAvailable, string name, bool preRelease) CheckNewVersion()
        {
            bool isNewVersionAvailable = false;
            string name = null;
            bool preRelease = false;

            Version currentVersion = FfeAssembly.AssemblyVersion;

            List<(string, bool, Version)> releases = new List<(string, bool, Version)>();

            try
            {
                foreach (JToken release in GitHubReleases)
                {
                    name = (string)release["name"];
                    preRelease = (bool)release["prerelease"];
                    Version version = Version.Parse(name);

                    releases.Add((name, preRelease, version));
                }

                var highestRelease = releases.SingleOrDefault(r => r.Item3 == releases.Max(mr => mr.Item3));

                if (highestRelease.Item3.CompareTo(currentVersion) > 0)
                {
                    isNewVersionAvailable = true;
                    name = highestRelease.Item1;
                    preRelease = highestRelease.Item2;
                }
            }
            catch (Exception)
            {
                Log.Error("New FFE Version could not be determined on basis of release name {@FfeVersion}", name ?? "null");
                throw;
            }

            return (isNewVersionAvailable, name, preRelease);
        }

        public static void CheckUpdate(bool showUpToDateMessage = true)
        {
            string title = "Update " + FfeAssembly.AssemblyTitle;

            try
            {
                var (isNewVersionAvailable, name, preRelease) = FfeUpdate.CheckNewVersion();
                if (isNewVersionAvailable)
                {
                    string strPreRelease = preRelease ? "pre-release " : null;
                    Forms.DialogResult dialogResult = Forms.MessageBox.Show($"A new FFE {strPreRelease}version {name} is available.\nDownload?", title, Forms.MessageBoxButtons.YesNo, Forms.MessageBoxIcon.Question);
                    if (dialogResult == Forms.DialogResult.Yes)
                    {
                        string downloadUrl = GetNewVersionDownloadUrl(name);
                        OpenLink(FfeSetting.Default.ChangelogUrl);
                        OpenLink(downloadUrl);
                    }
                }
                else if (showUpToDateMessage)
                {
                    Forms.MessageBox.Show("FFE is up to date.", title, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Information);
                }
            }
            catch (Exception)
            {
                Forms.DialogResult dialogResult = Forms.MessageBox.Show("An error occured.\nCheck and Download manually?", title, Forms.MessageBoxButtons.YesNo, Forms.MessageBoxIcon.Error);
                if (dialogResult == Forms.DialogResult.Yes)
                {
                    OpenLink(FfeSetting.Default.ReleasesUrl);
                }
            }
        }

        public static string GetNewVersionDownloadUrl(string releaseName)
        {
            string ffeXllFile = FfeSetting.Default.ReleaseFileName;
            string bitness = Environment.Is64BitProcess ? "x64" : "x32";
            var assets = GitHubReleases.Children().Single(r => r["name"].Value<string>().Equals(releaseName))["assets"];
            var downloadUrl = from asset in assets
                              where asset["name"].Value<string>().Equals(string.Format(ffeXllFile, bitness))
                              select asset;
            return downloadUrl.Single()["browser_download_url"].Value<string>();
        }

        private static void OpenLink(string url)
        {
            System.Diagnostics.Process.Start(url);
        }
    }
}