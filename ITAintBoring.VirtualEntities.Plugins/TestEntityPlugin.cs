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




        public SqlConnection getConnection(IPluginExecutionContext context)
        {
            SqlConnection result = null;
            if (context.SharedVariables.Contains("connectionstring"))
            {
                result = new SqlConnection((string)context.SharedVariables["connectionstring"]);
                try
                {
                    result.Open();
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException("Cannot open SQL concection");
                }
                return result;
            }
            else throw new InvalidPluginExecutionException("Connection string is not defined");
        }

        public TestEntityPlugin(string unsecureString, string secureString)
        {

            connectionString = secureString;
        }


        public void ParseConditionCollection(DataCollection<ConditionExpression> conditions, ref string keyword, ref Guid recordId)
        {
            foreach (ConditionExpression condition in conditions)
            {
                if (condition.Operator == ConditionOperator.Like && condition.Values.Count > 0)
                {
                    //This is how "search" works
                    keyword = (string)condition.Values[0];
                    break;
                }
                else if (condition.Operator == ConditionOperator.Equal)
                {
                    //throw new InvalidPluginExecutionException(condition.Values[0].ToString());
                    //This is to support $filter=<id_field> eq <value> sytax
                    //When this is called from the embedded canvas app
                    recordId = (Guid)condition.Values[0];
                    break;
                }
            }
        }

        public void Execute(IServiceProvider serviceProvider)
        {
            
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            if (context.Stage == 20)
            {
                context.SharedVariables["connectionstring"] = connectionString;
                return;
            }

            if (context.MessageName == "RetrieveMultiple")
            {
                
                if (context.Stage == 30)
                {
                    string keyword = null;
                    Guid recordId = Guid.Empty;
                    var qe = (QueryExpression)context.InputParameters["Query"];
                     if(qe.Criteria.Conditions.Count > 0)
                    {
                        //This is to support 
                        //ita_testvirtualentities?$filter=ita_testvirtualentityid+eq+...
                        ParseConditionCollection(qe.Criteria.Conditions, ref keyword, ref recordId);
                    }
                    else if(qe.Criteria.Filters.Count > 0)
                    {
                        //This is for the "search"
                        ParseConditionCollection(qe.Criteria.Filters[0].Conditions, ref keyword, ref recordId);
                    }
                    
                    if (keyword != null && keyword.StartsWith("[%]")) keyword = "%" + (keyword.Length > 2 ? keyword.Substring(3) : "");
                    loadEntities(context, keyword, recordId);
                    context.OutputParameters["BusinessEntityCollection"] = ec;
                }
            }
            else if(context.MessageName == "Retrieve")
            {
                loadEntities(context, null, context.PrimaryEntityId);
                context.OutputParameters["BusinessEntity"] = ec.Entities.ToList().Find(e => e.Id == context.PrimaryEntityId);
            }
        }


        public void loadEntities(IPluginExecutionContext context, string keyword, Guid id)
        {
            

            using (var con = getConnection(context))
            {

                System.Data.SqlClient.SqlCommand command = new System.Data.SqlClient.SqlCommand();
                
                command.Connection = con;
                if (id != Guid.Empty)
                {
                    command.CommandText = "SELECT Id, FirstName, LastName, Email FROM ITAExternalContact " +
                       "WHERE Id = @Id";
                    command.Parameters.Add("@Id", SqlDbType.UniqueIdentifier).Value = id;
                    
                }
                else  if(keyword != null)
                {
                    command.CommandText = "SELECT TOP 3 Id, FirstName, LastName, Email FROM ITAExternalContact " +
                         "WHERE FirstName like @Keyword OR LastName like @Keyword OR Email like @Keyword";
                    command.Parameters.Add("@Keyword", SqlDbType.NVarChar).Value = keyword;
                }
                else
                {
                    //When there are no search parameters
                    command.CommandText = "SELECT TOP 3 Id, FirstName, LastName, Email FROM ITAExternalContact ";
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
