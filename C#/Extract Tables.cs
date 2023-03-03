using System;
using ADO = Microsoft.AnalysisServices.AdomdClient;
using System.Data;
using System.Text; // string builder
using System.IO;

namespace Export_Power_BI_Tables
{
    class Program
    {
        enum SSASType
        {
            PbiDesktop,
            PbiService,
            SsasOnPremise
        }
        
        static void Main(string[] args)
        {
            string server = "localhost:52966"; // xmla endpoint in case of Power BI Service, server name in case of SSAS on premise
            string database = "Contoso 500K"; // dataset name or database name, in case of PBI desktop no need to specify
            string userName = "abcdef.onmicrosoft.com"; // windows or PBI Service Login name
            string userPassword = "****************";// windows or PBI Service Login password

            extractTableWithDAXQuery(server, database, userName, userPassword, SSASType.PbiDesktop);
        }

        static void extractTableWithDAXQuery(string server, string database, string userName, string userPassword, SSASType ssasType = SSASType.PbiDesktop)
        {
            string connString = "";

            switch (ssasType)
            {
                case SSASType.PbiDesktop:
                    connString = $@"DataSource={server};";
                    break;
                case SSASType.PbiService:
                    connString = $@"DataSource={server};Initial Catalog={database};User ID={userName};Password={userPassword};";
                    break;
                case SSASType.SsasOnPremise:
                    connString = $@"DataSource={server};Initial Catalog={database};User ID={userName};Password={userPassword};";
                    break;
            }

            string daxQuery = @"
                DEFINE 
                    MEASURE Sales[Sales Amount] = 
                        SUMX ( Sales, Sales[Quantity] * Sales[Net Price] )

                EVALUATE 
                    SUMMARIZECOLUMNS(
                        Dates[Year],
                        Dates[Monthw],
                        Products[Color],
                        ""Sales Amount"", [Sales Amount]
                    )
                    ORDER BY
                        Dates[Year] ASC,
                        [Sales Amount] DESC
            ";

            ADO.AdomdDataAdapter dataadapter = new(daxQuery, connString);
            DataTable table = new DataTable();

            try { dataadapter.Fill(table); }
            catch (Exception e) { Console.WriteLine(e.Message); }

            DataRow[] dataRowArray = new DataRow[table.Rows.Count];
            table.Rows.CopyTo(dataRowArray, 0);

            StringBuilder sb = new();

            string col = "";
            foreach (var c in table.Columns)
            {
                col += (c.ToString()).Split('[', ']')[1] + ",";
            }
            sb.AppendLine(col);

            foreach (DataRow row in table.Rows)
            {
                string line = "";
                for (int i = 0; i < row.ItemArray.Length; i++)
                {
                    line += row[i] + ",";
                }
                line = line.TrimEnd(',');
                sb.AppendLine(line);
            }

            string filePath = @"C:\Users\antsharma\Downloads\Power BI Tables.csv";
            File.WriteAllText(filePath, sb.ToString());
        }
    }
}
