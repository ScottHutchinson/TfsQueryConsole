// https://stackoverflow.com/a/10296612/5652483

using System;
using System.Configuration;
using System.IO;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Server;

namespace TfsQueryConsole {
    class Program {

        // http://blogs.microsoft.co.il/shair/2009/01/13/tfs-api-part-3-get-project-list-using-icommonstructureservice/
        static ProjectInfo[] GetDefaultProjectInfo(TfsTeamProjectCollection tfs) {
            // Create ICommonStructureService object that will take TFS Structure Service.
            ICommonStructureService structureService = (ICommonStructureService)tfs.GetService(typeof(ICommonStructureService));
            // Use ListAllProjects method to get all Team Projects in the TFS.
            ProjectInfo[] projects = structureService.ListAllProjects();
            return projects;
        }

        /* This function queries the selected TFS project for recent changesets and writes
         * the results, including full, untruncated comments in a .csv file.
        */
        static void ExportChangesetsWithFullComments(string author, bool sortAscending) {
            const int maxChangesets = 10000;
            //Uri defaultCollectionURI = new Uri("http://ici-ox-tfs:8080/tfs/DARTCollection");
            Uri defaultCollectionURI = new Uri(ConfigurationManager.AppSettings["defaultCollectionURI"]);
            TeamProjectPicker tpp = new TeamProjectPicker(TeamProjectPickerMode.SingleProject, false);
            var defaultCollection = new TfsTeamProjectCollection(defaultCollectionURI);
            tpp.SelectedTeamProjectCollection = defaultCollection;
            tpp.SelectedProjects = GetDefaultProjectInfo(defaultCollection);
            var dlgResult = tpp.ShowDialog();
            if (dlgResult != System.Windows.Forms.DialogResult.OK) { return; }
            string input = Microsoft.VisualBasic.Interaction.InputBox("Enter a positive integer:", "Number of Days Past", "30", -1, -1);
            if (string.IsNullOrEmpty(input)) { return; }
            int numDays = 0;
            bool isSuccessful = Int32.TryParse(input, out numDays);
            if (!isSuccessful || numDays <= 0) { return; }
            DateTime earliest = DateTime.Today.AddDays(-numDays);
            var tpc = tpp.SelectedTeamProjectCollection;
            var projectName = tpp.SelectedProjects[0].Name;
            VersionControlServer versionControl = tpc.GetService<VersionControlServer>();
            var tp = versionControl.GetTeamProject(projectName);
            var path = tp.ServerItem;
            var queryHistoryParameters = new QueryHistoryParameters(path, RecursionType.Full) {
                ItemVersion = VersionSpec.Latest,
                DeletionId = 0,
                Author = author,
                VersionStart = new DateVersionSpec(earliest),
                VersionEnd = null,
                MaxResults = maxChangesets,
                IncludeChanges = false,
                SlotMode = true,
                IncludeDownloadInfo = false,
                SortAscending = sortAscending
            };

            var q = versionControl.QueryHistory(queryHistoryParameters);
            int rowsReturned = 0;
            // Write the query results to a new file named "[ProjectName] Changesets with full comments [timestamp].csv".
            var outputFileName = string.Format("{0} Changesets with full comments {1:yyyy-MM-dd hhmm tt}.csv", projectName, DateTime.Now);
            string defaultOutputFolder = ConfigurationManager.AppSettings["outputFolder"];
            var outputPath = Path.Combine(defaultOutputFolder, outputFileName);
            using (StreamWriter outputFile = new StreamWriter(outputPath)) {
                outputFile.WriteLine("Changeset,User,Date,Comment");
                foreach (Changeset cs in q) {
                    var comment = cs.Comment;
                    // Replace double quotes with single quotes to make it simpler for Excel
                    // to parse the .csv file.
                    comment = comment.Replace("\"", "'");
                    // Wrap the comment in double quotes.
                    outputFile.WriteLine(string.Format(@"{0},""{1}"",{2},""{3}""", cs.ChangesetId, cs.OwnerDisplayName, cs.CreationDate, comment));
                    ++rowsReturned;
                }
            }
            Console.WriteLine("{0:N0} rows written to {1}\r\n", rowsReturned, outputPath);
            Console.WriteLine("Press any key to exit...", outputPath);
            Console.ReadKey(); // Wait for user to press a key
        }

        static void Main(string[] args) {
            // For now, assume that arg 0 is the name of the Author to search for, and arg 1 is the sortAscending flag.
            // TODO: Support multiple args in whatever order.
            // Example: TfsQueryConsole.exe -author:ICI\HutchinsonS -sortAscending:true
            int colonPos = 0;
            string author = null;
            if (args.Length > 0) {
                var authorArg = args[0];
                colonPos = authorArg.IndexOf(':');
                author = authorArg.Substring(colonPos + 1);
            }
            bool sortAscending = false;
            if (args.Length > 1) {
                var sortArg = args[1];
                colonPos = sortArg.IndexOf(':');
                bool.TryParse(sortArg.Substring(colonPos + 1), out sortAscending);
            }
            ExportChangesetsWithFullComments(author, sortAscending);
        }
    }
}
