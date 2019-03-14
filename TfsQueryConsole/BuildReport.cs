// https://blog.briancmoses.com/2012/02/viewing-all-queued-builds-in-tfs-part-2.html?r=related

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.Build.WebApi;
using Microsoft.VisualStudio.Services.Client;
using Microsoft.VisualStudio.Services.WebApi;

namespace TfsQueryConsole {
    class BuildReport {

        public static async Task<string> ReportBuildInfo() {
            Uri collectionURI = new Uri(ConfigurationManager.AppSettings["defaultCollectionURI"]);
            Console.WriteLine("\nTFS Build Queue");
            Console.WriteLine("===============\n");
            Console.WriteLine("Connecting to: " + collectionURI + " and querying build controllers...");
            var teamProjectName = "DART";
            var targetBuildName = "NG-DART-VS2012";
            VssConnection connection = new VssConnection(collectionURI, new VssClientCredentials());
            var buildserver = connection.GetClient<BuildHttpClient>();
            var builds = await buildserver.GetBuildsAsync(
                statusFilter: BuildStatus.Completed,
                project: teamProjectName);
            var targetedBuilds = builds
                .Where(definition => definition.Definition.Name.Contains(targetBuildName))
                .OrderBy(b => b.FinishTime)
                .ToList();

            return ProcessBuilds(targetedBuilds);
        }

        private static string ProcessBuilds(List<Build> targetedBuilds) {
            var pathString = $@"BuildResults.csv";
            var stringBuilder = new StringBuilder();
            stringBuilder.AppendLine("Name, Date, Passed, Build Time, Queue Time");
            foreach (var build in targetedBuilds) {
                var buildInfo = new BuildInfo {
                    Name = build.BuildNumber,
                    Passed = build.Result != null && build.Result.Value == BuildResult.Succeeded,
                    FinishTime = build.FinishTime.Value,
                    Started = build.StartTime.Value,
                    QueueStartTime = build.QueueTime.Value
                };

                stringBuilder.AppendLine(buildInfo.ToString());
            }

            File.WriteAllText(pathString, stringBuilder.ToString());
            return String.Format("Results written to {0}\r\n", Path.GetFullPath(pathString));
        }
    }

    internal class BuildInfo {
        public string Name { private get; set; }
        public bool Passed { private get; set; }
        public DateTime Started { private get; set; }
        public DateTime FinishTime { private get; set; }
        public DateTime QueueStartTime { private get; set; }

        private string InQueueTimeMinutes => this.CalculateTimeInQueue();

        /// <summary>
        /// The time the build took to run.
        /// </summary>
        public string TimeRanInMinutes => this.CalculateTimeRan();

        private string CalculateTimeRan() {
            var diff = this.FinishTime - this.Started;
            return (diff.TotalSeconds / 60).ToString(CultureInfo.CurrentCulture);
        }

        private string CalculateTimeInQueue() {
            var diff = this.Started - this.QueueStartTime;
            return (diff.TotalSeconds / 60).ToString(CultureInfo.CurrentCulture);
        }

        public override string ToString() {
            return $"{this.Name}, {this.Started.ToLocalTime():g}, {this.Passed}, {this.TimeRanInMinutes}, {this.InQueueTimeMinutes}";
        }
    }
}
