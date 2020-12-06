using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ITAintBoring.VirtualEntities.Plugins
{
    public class TestEntityPlugin : IPlugin
    {
        EntityCollection ec = new EntityCollection();
        string connectionString = "";




        public SqlConnection getConnection()
        {
            SqlConnection result = new SqlConnection(connectionString);
            
            return result;
        }
        
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            if (context.MessageName == "RetrieveMultiple")
            {
                
                if (context.Stage == 30)
                {
                    string keyword = "%";
                    var qe = (QueryExpression)context.InputParameters["Query"];
                    var filters = qe.Criteria.Filters;
                    if (filters.Count > 0)
                    {
                        foreach (FilterExpression filter in filters)
                        {
                            foreach(ConditionExpression condition in filter.Conditions)
                                if (condition.Operator == ConditionOperator.Like && condition.Values.Count > 0)
                                {
                                    keyword = (string)condition.Values[0];
                                    break;
                                }
                            break;
                        }
                    }
                    if (keyword.StartsWith("[%]")) keyword = "%" + (keyword.Length > 2 ? keyword.Substring(3) : "");
                    loadEntities(keyword, Guid.Empty);
                    context.OutputParameters["BusinessEntityCollection"] = ec;
                }
            }
            else if(context.MessageName == "Retrieve")
            {
                loadEntities("%", Guid.Empty);
                context.OutputParameters["BusinessEntity"] = ec.Entities.ToList().Find(e => e.Id == context.PrimaryEntityId);
            }
        }


        public void loadEntities(string keyword, Guid id)
        {
            using (var con = getConnection())
            {

                System.Data.SqlClient.SqlCommand command = new System.Data.SqlClient.SqlCommand();

                con.Open();
                command.Connection = con;
                if (id == Guid.Empty)
                {
                    command.CommandText = "SELECT TOP 3 Id, FirstName, LastName, Email FROM ITAExternalContact " +
                        "WHERE FirstName like @Keyword OR LastName like @Keyword OR Email like @Keyword";
                    command.Parameters.Add("@Keyword", SqlDbType.NVarChar).Value = keyword;
                }
                else
                {
                    command.CommandText = "SELECT Id, FirstName, LastName, Email FROM ITAExternalContact " +
                        "WHERE Id = @Id";
                    command.Parameters.Add("@Id", SqlDbType.UniqueIdentifier).Value = id;
                }
                

                SqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        
                        Entity e = new Entity("ita_testvirtualentity");
                        e.Attributes.Add("ita_testvirtualentityid", reader.GetGuid(0));
                        string firstName = reader.GetString(1);
                        string lastName = reader.GetString(2);
                        string email = reader.GetString(3);
                        e["ita_name"] = firstName + " " + lastName;
                        e["ita_firstname"] = firstName;
                        e["ita_lastname"] = lastName;
                        e["ita_email"] = email;
                        e.Id = (Guid)e.Attributes["ita_testvirtualentityid"];
                        ec.Entities.Add(e);
                    }
                }
                con.Close();
            }
        }
    }
}
