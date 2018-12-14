using ClosedXML.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace TFSPermissions
{
    class Program
    {
        static void Main(string[] args)
        {
            var url = "http://alm.eurofins.local/tfs/Eurofinscollection/_apis/projects?api-version=2.0";
            int iterator = 0;
            NetworkCredential networkCredential = new NetworkCredential("olx7", "Eurobgl@19");
            WebClient client = new WebClient();
            client.Credentials = networkCredential;
            var projectNamesJson = client.DownloadString(url);

            //List<Team> teams = new List<Team>();
            List<Project> projects = new List<Project>();
            List<Member> members = new List<Member>();
            List<Group> groups = new List<Group>();
            List<string> projectNames = ConverJsonToList(projectNamesJson, "value", "name");


            foreach (var projectname in projectNames)
            {
                iterator++;
                int teamIterator = 0;
                var teamNameURL = string.Format("http://alm.eurofins.local/tfs/Eurofinscollection/_apis/projects/{0}/teams?api-version=2.0", projectname);
                var teamNamesJson = client.DownloadString(teamNameURL);
                List<string> teamNames = ConverJsonToList(teamNamesJson, "value", "name");
                List<Team> teams = new List<Team>();

                foreach (var teamname in teamNames)
                {
                    teamIterator++;
                    var memberURL = string.Format("http://alm.eurofins.local/tfs/Eurofinscollection/_apis/projects/{0}/teams/{1}/members?api-version=2.0", projectname, teamname);
                    var memberNamesJson = client.DownloadString(memberURL);
                    List<string> memberNames = ConverJsonToList(memberNamesJson, "value", "displayName");
                    List<string> memberIds = ConverJsonToList(memberNamesJson, "value", "id");
                    Dictionary<string, string> memberDetails = GetUserDetails(memberNamesJson);
                    int memberIterator = 0;
                    foreach (var memberId in memberIds)
                    {
                        memberIterator++;
                        var permissionURL = string.Format("http://alm.eurofins.local/tfs/EurofinsCollection/_api/_identity/ReadGroupMembers?__v=5&scope={0}&readMembers=false", memberId);
                        var permissionGroupsJson = client.DownloadString(permissionURL);
                        List<string> PermissionGroup = ConverJsonToList(permissionGroupsJson, "identities", "DisplayName");
                        int permissionIteraor = 0;
                        foreach (var permission in PermissionGroup)
                        {
                            permissionIteraor++;
                            groups.Add(new Group()
                            {
                                //Id = permissionIteraor,
                                Permission = permission
                            });

                        }
                        members.Add(new Member()
                        {
                           // Id = memberIterator,
                            name = memberDetails[memberId],
                            Permissions = groups
                        });
                    }
                    teams.Add(new Team()
                    {
                       /// Id = teamIterator,
                        name = teamname,
                        Members = members
                    });
                }
                projects.Add(new Project()
                {
                    Id = iterator,
                    Name = projectname,
                    TeamNames = teams
                });
            }
            var jsonResult = JsonConvert.SerializeObject(projects);
            DataTable dataTableProjects = (DataTable)JsonConvert.DeserializeObject(jsonResult, (typeof(DataTable)));
            XLWorkbook wb = new XLWorkbook();            
            wb.Worksheets.Add(dataTableProjects, "WorksheetName");
            MemoryStream fs = new MemoryStream();
            wb.SaveAs(fs);
            fs.Position = 0;
        }

        public static List<string> ConverJsonToList(string json, string root, string filter)
        {
            List<string> projectNames = new List<string>();
            var projectNamesObject = JsonConvert.DeserializeObject(json);
            JArray projectNamesArray = new JArray();
            projectNamesArray.Add(projectNamesObject);
            for (int i = 0; i < projectNamesArray[0][root].Count(); i++)
            {
                projectNames.Add(projectNamesArray[0][root][i][filter].ToString());
            }
            return projectNames;
        }
        public static Dictionary<string,string> GetUserDetails(string json)
        {
            Dictionary<string, string> Details = new Dictionary<string, string>();
            var projectNamesObject = JsonConvert.DeserializeObject(json);
            JArray projectNamesArray = new JArray();
            projectNamesArray.Add(projectNamesObject);
            for (int i = 0; i < projectNamesArray[0]["value"].Count(); i++)
            {
                Details.Add(projectNamesArray[0]["value"][i]["id"].ToString(),projectNamesArray[0]["value"][i]["displayName"].ToString());
            }
            return Details;
        }

        public static string table_to_csv(DataTable table)
        {
            string file = "";

            foreach (DataColumn col in table.Columns)
                file = string.Concat(file, col.ColumnName, ",");

            file = file.Remove(file.LastIndexOf(','), 1);
            file = string.Concat(file, "\r\n");

            foreach (DataRow row in table.Rows)
            {
                foreach (object item in row.ItemArray)
                    file = string.Concat(file, item.ToString(), ",");

                file = file.Remove(file.LastIndexOf(','), 1);
                file = string.Concat(file, "\r\n");
            }

            return file;
        }
    }
}
