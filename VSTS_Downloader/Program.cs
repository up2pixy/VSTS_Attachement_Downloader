using CommandLine;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Client;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace VSTS_Downloader
{
    class Program
    {
        private const string clientId = "{Your client id created from Azure Portal}";
        private const string replyUri = "{Your replyUri defined in Azure Portal}";
        private const string VSTSResourceId = "499b84ac-1321-427f-aa17-267ca6975798"; //This is the resource id for VSTS in Azure (No need to change)


        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
                .WithParsed<Options>(opts => RunOptionsAndReturnExitCode(opts))
                .WithNotParsed<Options>((errs) => HandleParseError(errs));
        }

        private static void HandleParseError(IEnumerable<Error> errs)
        {
        }

        private static void RunOptionsAndReturnExitCode(Options opts)
        {
            Console.WriteLine($"Authorizing to VSTS...");
            try
            {
                VssConnection connection = new VssConnection(new Uri(opts.CollectionUrl), new VssClientCredentials(false));

                WorkItemTrackingHttpClient witClient = connection.GetClient<WorkItemTrackingHttpClient>();
                WorkItemQueryResult queryResults = witClient.QueryByIdAsync(opts.QueryId).Result;

                if (queryResults == null)
                {
                    Console.WriteLine("Query result is null.");
                    return;
                }

                var idList = new List<int>();
                if (queryResults.QueryType == QueryType.Flat)
                {
                    if (queryResults.WorkItems == null || queryResults.WorkItems.Count() == 0)
                    {
                        Console.WriteLine("Query did not find any results");
                        return;
                    }
                    foreach (var item in queryResults.WorkItems)
                    {
                        idList.Add(item.Id);
                    }
                }
                else
                {
                    if (queryResults.WorkItemRelations == null || queryResults.WorkItemRelations.Count() == 0)
                    {
                        Console.WriteLine("Query did not find any results");
                        return;
                    }
                    foreach (var item in queryResults.WorkItemRelations)
                    {
                        idList.Add(item.Target.Id);
                    }
                }

                foreach (int id in idList)
                {
                    var attachmentString = opts.Latest ? "the latest attachment" : "all attachments";
                    Console.Write($"Downloading {attachmentString} for {id}... ");

                    var workitem = witClient.GetWorkItemAsync(id, null, null, WorkItemExpand.Relations, null).Result;

                    if (workitem.Relations == null || workitem.Relations.Count == 0 || workitem.Relations.All(r => r.Rel != "AttachedFile"))
                    {
                        Console.WriteLine("No attachment");
                        continue;
                    }

                    foreach (var relation in workitem.Relations.Where(r => r.Rel == "AttachedFile").OrderByDescending(r => (DateTime)r.Attributes.GetValueOrDefault("authorizedDate")))
                    {
                        //Get the guid from the end of the url                   
                        string attachmentId = relation.Url.ToString().Split('/').Last();
                        var filename = relation.Attributes.GetValueOrDefault("name").ToString();
                        Stream attachmentStream = witClient.GetAttachmentContentAsync(new Guid(attachmentId)).Result;
                        var fileMode = opts.Overwrite ? FileMode.Create : FileMode.Create;
                        using (FileStream writeStream = new FileStream(CreateOutputFilePath(opts, id, filename), fileMode, FileAccess.ReadWrite))
                        {
                            attachmentStream.CopyTo(writeStream);
                        }
                        if (opts.Latest)
                        {
                            break;
                        }
                    }
                    Console.WriteLine("Done");
                }

                Console.WriteLine("All Done!");
            }
            catch (Exception ex)
            {
                Exception innerEx = ex;
                while (innerEx.InnerException != null)
                {
                    innerEx = innerEx.InnerException;
                }
                Console.WriteLine(innerEx.Message);
            }
        }

        private static string CreateOutputFilePath(Options opts, int id, string targetFileName)
        {
            var outputFolder = opts.OutputFolder.TrimEnd('\\');

            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            string filepath = string.Empty;

            if (opts.Flat)
            {
                filepath = outputFolder + "\\" + targetFileName;
            }
            else
            {
                var subfolder = outputFolder + "\\" + id;
                if (!Directory.Exists(subfolder))
                {
                    Directory.CreateDirectory(subfolder);
                }
                filepath = subfolder + "\\" + targetFileName;
            }
            return filepath;
        }

        private class Options
        {
            [Option(
                's',
                "server",
                Required = true,
                HelpText = "VSTS server address")]
            public string CollectionUrl { get; set; }


            [Option(
                'q',
                "query",
                Required = true,
                HelpText = "VSTS query Id")]
            public Guid QueryId { get; set; }

            [Option(
                'o',
                "output",
                Required = true,
                HelpText = "Output folder path")]
            public string OutputFolder { get; set; }

            [Option(
              Default = false,
              HelpText = "Download attachments to same folder")]
            public bool Flat { get; set; }

            [Option(
                Default = false,
                HelpText = "Download the latest attachment for each work item instead of all attachments")]
            public bool Latest { get; set; }

            [Option(
                Default = false,
                HelpText = "Overwrite file if exist")]
            public bool Overwrite { get; set; }
        }
    }
}
