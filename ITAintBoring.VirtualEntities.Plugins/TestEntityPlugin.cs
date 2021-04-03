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

            string conString = connectionString;
            if (context.SharedVariables.Contains("connectionstring")) conString = (string)context.SharedVariables["connectionstring"];
            SqlConnection result = null;
            try
            {
                result = new SqlConnection(conString);
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
            catch(Exception ex)
            {
                throw new InvalidPluginExecutionException("Error connecting to SQL: " + ex.Message);
            }
        }

        public TestEntityPlugin(string unsecureString, string secureString)
        {
            if(!String.IsNullOrEmpty(secureString)) connectionString = secureString;
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

        public void ExecuteRetrieveMultiple(IPluginExecutionContext context)
        {
            string keyword = null;
            Guid recordId = Guid.Empty;
            var qe = (QueryExpression)context.InputParameters["Query"];
            if (qe.Criteria.Conditions.Count > 0)
            {
                //This is to support 
                //ita_testvirtualentities?$filter=ita_testvirtualentityid+eq+...
                ParseConditionCollection(qe.Criteria.Conditions, ref keyword, ref recordId);
            }
            else if (qe.Criteria.Filters.Count > 0)
            {
                //This is for the "search"
                ParseConditionCollection(qe.Criteria.Filters[0].Conditions, ref keyword, ref recordId);
            }

            if (keyword != null && keyword.StartsWith("[%]")) keyword = "%" + (keyword.Length > 2 ? keyword.Substring(3) : "");
            loadEntities(context, keyword, recordId);
            context.OutputParameters["BusinessEntityCollection"] = ec;
        }

        public void ExecuteRetrieve(IPluginExecutionContext context)
        {
            loadEntities(context, null, context.PrimaryEntityId);
            context.OutputParameters["BusinessEntity"] = ec.Entities.ToList().Find(e => e.Id == context.PrimaryEntityId);
        }

        public void ExecuteCreate(IPluginExecutionContext context)
        {
            using (var con = getConnection(context))
            {
                
                Entity target = (Entity)context.InputParameters["Target"];
                target.Id = Guid.NewGuid();
                System.Data.SqlClient.SqlCommand command = new System.Data.SqlClient.SqlCommand();
                command.Connection = con;
                List<string> valueList = new List<string>();
                List<string> fieldList = new List<string>();
                string commandText = "INSERT INTO ITAExternalContact ({0}) VALUES ({1})";
                
                fieldList.Add("Id");
                valueList.Add("@Id");
                
                command.Parameters.Add("@Id", SqlDbType.UniqueIdentifier).Value = target.Id;
                if (target.Contains("ita_firstname") && target["ita_firstname"] != null)
                {
                    fieldList.Add("FirstName");
                    valueList.Add("@FirstName");
                    command.Parameters.Add("@FirstName", SqlDbType.NVarChar).Value = target["ita_firstname"];
                }
                if (target.Contains("ita_lastname") && target["ita_lastname"] != null)
                {
                    fieldList.Add("LastName");
                    valueList.Add("@LastName");
                    command.Parameters.Add("@LastName", SqlDbType.NVarChar).Value = target["ita_lastname"];
                }
                if (target.Contains("ita_email") && target["ita_email"] != null)
                {
                    fieldList.Add("Email");
                    valueList.Add("@Email");
                    command.Parameters.Add("@Email", SqlDbType.NVarChar).Value = target["ita_email"];
                }
                
                command.CommandText = string.Format(commandText, string.Join(",", fieldList), string.Join(",", valueList));
                command.ExecuteNonQuery();

                context.OutputParameters["id"] = target.Id;
            }
        }

        public void ExecuteUpdate(IPluginExecutionContext context)
        {
            using (var con = getConnection(context))
            {
                Entity target = (Entity)context.InputParameters["Target"];
                System.Data.SqlClient.SqlCommand command = new System.Data.SqlClient.SqlCommand();
                command.Connection = con;
                string commandText = "UPDATE ITAExternalContact SET {0} WHERE Id=@Id";
                List<string> setList = new List<string>();
                command.Parameters.Add("@Id", SqlDbType.UniqueIdentifier).Value = target.Id;
                if (target.Contains("ita_firstname"))
                {
                    setList.Add("FirstName=@FirstName");
                    command.Parameters.Add("@FirstName", SqlDbType.NVarChar).Value = target["ita_firstname"];
                }
                if (target.Contains("ita_lastname"))
                {
                    setList.Add("LastName=@LastName");
                    command.Parameters.Add("@LastName", SqlDbType.NVarChar).Value = target["ita_lastname"];
                }
                if (target.Contains("ita_email"))
                {
                    setList.Add("Email=@Email");
                    command.Parameters.Add("@Email", SqlDbType.NVarChar).Value = target["ita_email"];
                }
                command.CommandText = string.Format(commandText, string.Join(",", setList));
                command.ExecuteNonQuery();
            }
        }

        public void ExecuteDelete(IPluginExecutionContext context)
        {
            using (var con = getConnection(context))
            {
                EntityReference target = (EntityReference)context.InputParameters["Target"];
                System.Data.SqlClient.SqlCommand command = new System.Data.SqlClient.SqlCommand();
                command.Connection = con;
                command.CommandText = "DELETE FROM ITAExternalContact " +
                             "WHERE Id = @Id";
                command.Parameters.Add("@Id", SqlDbType.UniqueIdentifier).Value = target.Id;
                command.ExecuteNonQuery();
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

            if (context.Stage == 30)
            {
                if (context.MessageName == "RetrieveMultiple")
                {
                    ExecuteRetrieveMultiple(context);
                }
                else if (context.MessageName == "Retrieve")
                {
                    ExecuteRetrieve(context);
                }
                else if (context.MessageName == "Create")
                {
                    ExecuteCreate(context);
                }
                else if (context.MessageName == "Update")
                {
                    ExecuteUpdate(context);
                }
                else if (context.MessageName == "Delete")
                {
                    ExecuteDelete(context);
                }
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
                    command.CommandText = "SELECT TOP 30 Id, FirstName, LastName, Email FROM ITAExternalContact " +
                         "WHERE FirstName like @Keyword OR LastName like @Keyword OR Email like @Keyword";
                    command.Parameters.Add("@Keyword", SqlDbType.NVarChar).Value = keyword;
                }
                else
                {
                    //When there are no search parameters
                    command.CommandText = "SELECT TOP 30 Id, FirstName, LastName, Email FROM ITAExternalContact ";
                }
               


                SqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        
                        Entity e = new Entity(context.PrimaryEntityName);
                        e.Attributes.Add(context.PrimaryEntityName+"id", reader.GetGuid(0));
                        string firstName = reader.IsDBNull(1) ? null : reader.GetString(1);
                        string lastName = reader.IsDBNull(2) ? null : reader.GetString(2);
                        string email = reader.IsDBNull(3) ? null : reader.GetString(3);
                        e["ita_name"] = firstName + " " + lastName;
                        e["ita_firstname"] = firstName;
                        e["ita_lastname"] = lastName;
                        e["ita_email"] = email;
                        e.Id = (Guid)e.Attributes[context.PrimaryEntityName + "id"];
                        ec.Entities.Add(e);
                    }
                }
                con.Close();
            }
        }
    }
}
