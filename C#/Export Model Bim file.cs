using Tabular = Microsoft.AnalysisServices.Tabular;
using JSON = Newtonsoft.Json;
using System.IO;

namespace Extract_Model_BIM_File
{
    class Program
    {
        enum SSASType
        {
            PbiDesktop,
            PbiService,
            SsasOnPremise
        }

        static Tabular.Server server = null;
        static Tabular.Database database = null;
        
        static void Main(string[] args)
        {
            string server = "localhost:52966"; // xmla endpoint in case of Power BI Service, server name in case of SSAS on premise
            string database = "Contoso 500K"; // dataset name or database name, in case of PBI desktop no need to specify
            string userName = "abcdef.onmicrosoft.com"; // windows or PBI Service Login name, in case of PBI desktop no need to specify
            string userPassword = "****************";// windows or PBI Service Login password, in case of PBI desktop no need to specify

            ConnectToServer(serverName, databaseName,userName, userPassword, SSASType.PbiService);
            GetMetadata();
            server.Disconnect();
        }

        static void GetMetadata()
        {
            var script = Tabular.JsonScripter.ScriptCreate(database);
            dynamic jsonObj = JSON.JsonConvert.DeserializeObject(script);
            jsonObj = jsonObj["create"]["database"];

            string output = JSON.JsonConvert.SerializeObject(jsonObj, JSON.Formatting.Indented);
            File.WriteAllText(@"C:\Users\antsharma\Downloads\model.bim", output);
        }
        
        static Tabular.Database ConnectToServer(
            string serverName, 
            string databaseName, 
            string userName, 
            string userPassword, 
            SSASType ssasType = SSASType.PbiDesktop)
        {
            string connString = "";

            switch (ssasType)
            {
                case SSASType.PbiDesktop:
                    connString = $@"DataSource={serverName};";
                    break;
                case SSASType.PbiService:
                    connString = $@"DataSource={serverName};Initial Catalog={databaseName};User ID={userName};Password={userPassword};";
                    break;
                case SSASType.SsasOnPremise:
                    connString = $@"DataSource={serverName};Initial Catalog={databaseName};User ID={userName};Password={userPassword};";
                    break;
            }

            server = new();
            server.Connect(connString);

            switch (ssasType)
            {
                case SSASType.PbiDesktop:
                    database = server.Databases[0];
                    break;
                case SSASType.SsasOnPremise:
                    database = server.Databases.GetByName(databaseName);
                    break;
                case SSASType.PbiService:
                    database = server.Databases.GetByName(databaseName);
                    break;
            }

            return database;
        }
    }
}
