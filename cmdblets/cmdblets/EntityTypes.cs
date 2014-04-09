using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.ServiceModel;
using System.Text.RegularExpressions;
using System.ServiceModel.Channels;
using cmdblets.CMDBUILD;
namespace cmdblets
{
    public class EntityTypes : CMDBCmdletBase
    {
        public PSObject AdaptRelationObject(Cmdlet myCmdlet, relation[] relObjects,string domain)
        {
            var promotedObject = new PSObject(relObjects);
            promotedObject.TypeNames.Insert(1, relObjects.GetType().FullName);
            promotedObject.TypeNames[0] = String.Format(CultureInfo.CurrentCulture, "RelationObject#{0}",domain);
            return promotedObject;
        }

        public  PSObject AdaptCardObject(Cmdlet myCmdlet, card cardObject)
        {
            var promotedObject = new PSObject(cardObject);
            promotedObject.TypeNames.Insert(1, cardObject.GetType().FullName);
            promotedObject.TypeNames[0] = String.Format(CultureInfo.CurrentCulture, "CardObject#{0}",cardObject.className);
            // loop through the properties and promote them into the PSObject we're going to return
            foreach ( var p in cardObject.attributeList)
            {
                try
                {
                    promotedObject.Members.Add(new PSNoteProperty(p.name, p.value));
                }
                catch (ExtendedTypeSystemException ets)
                {
                    myCmdlet.WriteWarning(String.Format("The property '{0}' already exists, skipping.\nException: {1}", p.name, ets.Message));
                }
                catch (Exception e)
                {
                    myCmdlet.WriteError(new ErrorRecord(e, "Property", ErrorCategory.NotSpecified, p.name));
                }
            }
            return promotedObject;
        }

        public int GetCardbyCode(string referencedClassName, string code)
        {
            int returnvalue = 0;
            var myQuery = new query();
            var fil = new filter
            {
                @operator = "LIKE",
                name = "Code",
                value = new string[] { "%" + code + "%" }
            };
            myQuery.filter = fil;
            var cardfound = _clientconnection.getCardList(referencedClassName, null, myQuery, null, 1, 0, null, null);
            if (cardfound.cards != null)
            {
                var r = cardfound.cards.FirstOrDefault();
                returnvalue = r.id;
            }
            return returnvalue;
        }

        public void AssignNewValue(Cmdlet myCmdlet, attributeSchema prop, card so, object newValue)
        {
            if (newValue == null) { return; } // Hacky Workaround
            var lists = so.attributeList != null ? so.attributeList.ToList() : new List<attribute>();
          
            myCmdlet.WriteVerbose("Want to set " + prop.name + " to " + newValue.ToString());
                
         
            var dict = lists.ToDictionary(key => key.name, value => value);
            var as1 = new attribute{ name = prop.name};
            switch (prop.type)
            {
                case  "DATE":
                    string valtoset = newValue != null ? newValue.ToString() : null;
                    if (!string.IsNullOrEmpty(valtoset))
                    {
                        try
                        {
                            myCmdlet.WriteVerbose("DATE: Convert string to DateTimeObject");
                            DateTime xd = DateTime.Parse(valtoset.ToString(), CultureInfo.CurrentCulture);
                            as1.value = xd.ToString("yyyy-MM-ddThh:mm:ssZ");
                        }
                        catch (Exception e) { myCmdlet.WriteWarning(e.Message); }
                    }
                    break;
                case "TIMESTAMP":
                    string datetimetoset = newValue != null ? newValue.ToString() : null;
                    if (!string.IsNullOrEmpty(datetimetoset))
                    {
                        try
                        {
                            myCmdlet.WriteVerbose("TIMESTAMP: Convert string to DateTimeObject -> _" + datetimetoset.ToString() +"_");
                            DateTime xd = DateTime.Parse(datetimetoset.ToString(), CultureInfo.CurrentCulture);
                            as1.value = xd.ToString("yyyy-MM-ddThh:mm:ssZ");
                        }
                        catch (Exception e) { myCmdlet.WriteWarning(e.Message); }
                    }
                    break;
                case "REFERENCE":
                    if (prop.referencedIdClassSpecified && newValue.ToString().Length > 0)
                    {

                        int codeint = 0;
                        int.TryParse(newValue.ToString(), out codeint);
                        if (codeint == 0)
                        {
                            myCmdlet.WriteVerbose("Try it with Workaround Parent Class: " + prop.referencedClassName);
                            int tmpvalue = 0;
                            var parentclass = _clientconnection.getAttributeList(prop.referencedClassName);
                            foreach (attributeSchema attributeschema in parentclass)
                            {
                                if (attributeschema.name == as1.name)
                                {
                                    string newclassname = attributeschema.referencedClassName;
                                    myCmdlet.WriteVerbose("Try it with Fulltext search classname: " + newclassname);
                                    tmpvalue = GetCardbyCode(newclassname, newValue.ToString());
                                    if (tmpvalue == 0)
                                    {
                                        myCmdlet.WriteWarning("Reference Card not found: " + newValue.ToString());
                                        dict.Remove(as1.name);
                                    }
                                    else
                                    {
                                        myCmdlet.WriteVerbose("Card found: " + tmpvalue);
                                        newValue = tmpvalue;
                                    }

                                    break;
                                }
                            }

                            
                            if (tmpvalue == 0)
                            {
                                myCmdlet.WriteVerbose("Fulltext search ReferenceCard: " + newValue.ToString());
                                myCmdlet.WriteVerbose("Fulltext search classname: " + prop.referencedClassName);
                                tmpvalue = GetCardbyCode(prop.referencedClassName, newValue.ToString());
                                newValue = tmpvalue;
                            }
                            else
                            {
                                myCmdlet.WriteVerbose("Card found Set to Value: " + tmpvalue);
                                newValue = tmpvalue;
                            }
 
                        }
                        else
                        {

                            //myCmdlet.WriteVerbose("Check if ReferenceCard exists: " + codeint);
                            //var check = _clientconnection.getCard(prop.referencedClassName, codeint, null);
                            //if (check == null)
                            //{
                            //    myCmdlet.WriteWarning("Reference Card not found: " + codeint);
                            //    dict.Remove(as1.name);
                            //}
                        }
                        as1.value = newValue.ToString();

                    }
                    break;
                default:
                    as1.value = newValue != null ? newValue.ToString() : null;
                    break;

            }
            if (dict.ContainsKey(as1.name))
            {
                dict[as1.name] = as1;
            }
            else
            {
                dict.Add(as1.name, as1);
            }
            so.attributeList = dict.Values.ToArray();

        }
        // All this does is convert the PowerShell filter language to the criteria syntax
        public string ConvertFilterToCMDBuildOperator(string filter)
        {
            var OpToOp = new Dictionary<string, string>();
            Regex re;
            // "-gt","-ge","-lt","-le","-eq","-ne","-like","-notlike","-match","-notmatch"
            // Add -isnull and -isnotnull, even though they aren't PowerShell operators
            OpToOp.Add("-and", "and");
            OpToOp.Add("-or", "or");
            OpToOp.Add("-eq", "EQUALS");
            OpToOp.Add("-ne", "DIFFERENT");
            OpToOp.Add("-lt", "STRICT_MINOR");
            OpToOp.Add("-gt", "STRICT_MAJOR");
            OpToOp.Add("-le", "MINOR");
            OpToOp.Add("-ge", "MAJOR");
            OpToOp.Add("-like", "LIKE");
            OpToOp.Add("-notlike", "DONTCONTAINS");
            OpToOp.Add("-isnull", "NULL");
            OpToOp.Add("-isnotnull", "is not null");
            re = new Regex("\\*");
            filter = re.Replace(filter, "%");
            re = new Regex("\\?");
            filter = re.Replace(filter, "_");
            re = new Regex("\"");
            filter = re.Replace(filter, "'");
            foreach (string k in OpToOp.Keys)
            {
                re = new Regex(k, RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
                filter = re.Replace(filter, OpToOp[k]);
            }
            return filter;
        }
        public query ConvertFilterToQuery(string filter)
        {
            query myQuery = null;
            try
            {
                var re1 = @"^(?<Attribute>\w{1,})+\s(?<Operator>\-\w{1,})+\s(?<Value>(\S|\w|\d|\s){1,})+$";

                WriteVerbose("Original Filter: " + filter);
                var r = new Regex(re1, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                var m = r.Match(filter);
                if (m.Success)
                {
                    var var1 = m.Groups["Attribute"].ToString();
                    var var2 = m.Groups["Operator"].ToString();
                    var var3 = m.Groups["Value"].ToString().Replace("*","%");
                    string filterString = ConvertFilterToCMDBuildOperator(var2);
                    //filterString = FixUpPropertyNames(filterString, mpClass);
                    WriteVerbose("Property: " + var1);
                    WriteVerbose("Operator: " + filterString);
                    WriteVerbose("Value: " + var3);
                    myQuery = new query(); //(filterString, mpClass);
                    var fil = new filter
                    {
                        @operator = filterString,
                        name = var1,
                        value = new string[] { var3 }
                    };

                    myQuery.filter = fil;
                    WriteVerbose("Using " + filterString + " as query");
                    return myQuery;
                }
 
            }
            catch // This is non-catastrophic - it's our first attempt
            {
                WriteDebug("failed: " + filter);
            }

            return myQuery;

        }
    }


    [Cmdlet(VerbsCommon.Get, "CMDBMenuSchema", DefaultParameterSetName = "No")]
    public class GetCMDBMenuSchemaCommand : EntityTypes
    {
        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            try
            {
                var schema = _clientconnection.getMenuSchema();
                WriteObject(schema);


            }
            catch (Exception e)
            {

                WriteError(new ErrorRecord(e, "Unknown error", ErrorCategory.NotSpecified, ""));
            }
        }
    }

    [Cmdlet(VerbsCommon.Get, "CMDBCardMenuSchema", DefaultParameterSetName = "No")]
    public class GetCMDBCardMenuSchemaCommand : EntityTypes
    {
        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            try
            {
                var cardmenuschema = _clientconnection.getCardMenuSchema();
                WriteObject(cardmenuschema);


            }
            catch (Exception e)
            {

                WriteError(new ErrorRecord(e, "Unknown error", ErrorCategory.NotSpecified, ""));
            }
        }
    }

    [Cmdlet(VerbsCommon.Get, "CMDBFunctionList", DefaultParameterSetName = "No")]
    public class GetCMDBFunctionListCommand : EntityTypes
    {
        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            try
            {
                var functions = _clientconnection.getFunctionList();
                WriteObject(functions);


            }
            catch (Exception e)
            {

                WriteError(new ErrorRecord(e, "Unknown error", ErrorCategory.NotSpecified, ""));
            }
        }
    }

    [Cmdlet(VerbsCommon.Get, "CMDBUserInfo",DefaultParameterSetName = "No") ]
    public class GetCMDBUserInfoCommand : EntityTypes
    {
        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            try
            {
                var user = _clientconnection.getUserInfo();
                WriteObject(user);


            }
            catch (Exception e)
            {

                WriteError(new ErrorRecord(e, "Unknown error", ErrorCategory.NotSpecified, ""));
            }
        }
    }

    [Cmdlet(VerbsCommon.New, "CMDBRelationshipObject", DefaultParameterSetName = "No")]
    public class NewCMDBRelationshipObjectCommand : EntityTypes
    {
        // Parameters
        private string _name = "";
        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public string RelationName
        {
            get { return _name; }
            set { _name = value; }
        }
        [Parameter(Position = 1, Mandatory = true, ValueFromPipeline = true,ParameterSetName = "ById")]
        public int SourceId { get; set; }
        [Parameter(Position = 2, Mandatory = true, ValueFromPipeline = true, ParameterSetName = "ById")]
        public int TargetId { get; set; }

        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            try
            {

                var rel = new relation {domainName = _name, card1Id = SourceId, card2Id = TargetId};
                var rval = _clientconnection.createRelation(rel);
                WriteObject(rval);


            }
            catch (Exception e)
            {

                WriteError(new ErrorRecord(e, "Unknown error", ErrorCategory.NotSpecified, _name));
            }

        }      
    }
    [Cmdlet(VerbsCommon.Push, "CMDBFunction", DefaultParameterSetName = "No")]
    public class PushCMDBFunctionCommand : EntityTypes
    {
        // Parameters
        private string _name;
        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            try
            {
                var attributes = _clientconnection.callFunction(_name,null);
                foreach (var attribute in attributes)
                {
                    WriteObject(attribute);
                }
                

            }
            catch (Exception e)
            {

                WriteError(new ErrorRecord(e, "Unknown error", ErrorCategory.NotSpecified, _name));
            }

        }
    }
    [Cmdlet(VerbsCommon.Set, "CMDBLookup", DefaultParameterSetName = "No")]
    public class SetCMDBLookupCommand : EntityTypes
    {
        // Parameters
        private lookup _lookup;
        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public lookup Lookup
        {
            get { return _lookup; }
            set { _lookup = value; }
        }

        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            try
            {
                var up = _clientconnection.updateLookup(_lookup);
                WriteObject(up);


            }
            catch (Exception e)
            {

                WriteError(new ErrorRecord(e, "Unknown error", ErrorCategory.NotSpecified, _lookup));
            }

        }
    }
    [Cmdlet(VerbsCommon.Get, "CMDBRelationshipObjectHistory", DefaultParameterSetName = "No")]
    public class GetCMDBRelationshipObjectHistoryCommand : EntityTypes
    {
        // Parameters
        private relation _relation;
        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public relation RelationObject
        {
            get { return _relation; }
            set { _relation = value; }
        }

        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            try
            {
                var relationHist = _clientconnection.getRelationHistory(_relation);

                if (relationHist != null)
                {
                    foreach (var relation in relationHist)
                    {
                        WriteObject(relation);
                    }
 
                }


            }
            catch (Exception e)
            {

                WriteError(new ErrorRecord(e, "Unknown error", ErrorCategory.NotSpecified, _relation));
            }

        }
    }

    [Cmdlet(VerbsCommon.Remove, "CMDBRelationshipObject", DefaultParameterSetName = "No")]
    public class RemoveCMDBRelationshipObjectCommand : EntityTypes
    {
        // Parameters
        private relation _relation;
        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public relation RelationObject
        {
            get { return _relation; }
            set { _relation = value; }
        }

        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            try
            {
                var res = _clientconnection.deleteRelation(_relation);
                WriteObject(res);


            }
            catch (Exception e)
            {

                WriteError(new ErrorRecord(e, "Unknown error", ErrorCategory.NotSpecified, _relation));
            }

        }
    }
    [Cmdlet(VerbsCommon.Get, "CMDBRelationshipObject", DefaultParameterSetName = "No")]
    public class GetCMDBRelationshipObjectCommand : EntityTypes
    {
        // Parameters
        private string _name = "";
        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public string RelationName
        {
            get { return _name; }
            set { _name = value; }
        }
        [Parameter(Position = 1, Mandatory = false, ValueFromPipeline = true)]
        public int Id { get; set; }

        public  GetCMDBRelationshipObjectCommand()
        {
            Id = 0;
        }
        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            try
            {
                var relationList = _clientconnection.getRelationList(_name, null, Id);

                if (relationList != null)
                {
                    //WriteObject(relationList);
                    WriteObject(AdaptRelationObject(this, relationList, _name));
                }


            }
            catch (Exception e)
            {

                WriteError(new ErrorRecord(e, "Unknown error", ErrorCategory.NotSpecified, _name));
            }
           
        }
    }

    [Cmdlet(VerbsCommon.New, "CMDBUpdateCreateObject", DefaultParameterSetName = "No")]
    public class NewCMDBUpdateCreateObjectCommand : EntityTypes
    {
        // Parameters
        private string _name = "";
        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public string ClassName
        {
            get { return _name; }
            set { _name = value; }
        }
        [Parameter(Position = 1, Mandatory = true, ValueFromPipeline = true)]
        public Hashtable PropertyHashtable { get; set; }

        [Parameter]
        public SwitchParameter PassThru { get; set; }

        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            try
            {
              
                var resclass = _clientconnection.getAttributeList(_name);
                if (resclass != null)
                {
                    var o = new card { className = _name };

                    var ht = new Hashtable(StringComparer.OrdinalIgnoreCase);
                    foreach (var prop in resclass)
                    {
                        try
                        {
                            WriteVerbose("Property in Class: " + prop.name);
                            ht.Add(prop.name, prop);
                        }
                        catch (Exception e)
                        {
                            WriteError(new ErrorRecord(e,
                                                       "property '" + prop.name +
                                                       "' has already been added to collection",
                                                       ErrorCategory.InvalidOperation, prop));
                        }
                    }

                    // now go through the hashtable that has the values we we want to use and
                    // assigned them into the new values
                    foreach (string s in PropertyHashtable.Keys)
                    {
                        if (!ht.ContainsKey(s))
                        {
                            WriteError(new ErrorRecord(new SystemException(s), "property not found on object", ErrorCategory.NotSpecified, o));
                        }
                        else
                        {

                            var p = ht[s] as attributeSchema;
                            if (p != null)
                            {
                                WriteVerbose(p.name);
                                AssignNewValue(this,p, o, PropertyHashtable[s]);
                            }
                        }
                    }
                    //Create Card
                    var query = new query();
                    var f = new filter();
                    f.@operator = "EQUALS";
                    f.name = "Code";
                    f.value = new[] {ht["Code"].ToString()};
                    query.filter = f;

                    var searchcard = _clientconnection.getCardList(o.className, null, query, null, 1, 0, null, null);
                    int rcard;
                    if (searchcard != null)
                    {
                        var firstOrDefault = searchcard.cards.FirstOrDefault();
                        if (firstOrDefault != null) rcard = firstOrDefault.id;
                        //_clientconnection.updateCard()
                    }
                    else
                    {
                       rcard = _clientconnection.createCard(o); 
                    }
                    
                    if (PassThru)
                    {
                        var r = _clientconnection.getCard(_name, 0, null);
                        WriteObject(AdaptCardObject(this, r));
                    }
                }
            }
            catch (Exception e)
            {

                WriteError(new ErrorRecord(e, "Unknown error", ErrorCategory.NotSpecified, ""));
            }

        }

    }

    [Cmdlet(VerbsCommon.Get, "CMDBClass", DefaultParameterSetName = "No")]
    public class GetCMDBClassCommand : EntityTypes
    {
        // Parameters
        private string _name = "Class";
        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public string ClassName
        {
            get { return _name; }
            set { _name = value; }
        }

        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            try
            {
                var resclass = _clientconnection.getClassSchema(_name);
                
                if (resclass != null)
                {
                    WriteObject(resclass);  
                }

                
            }
            catch (Exception e)
            {

                WriteError(new ErrorRecord(e, "Unknown error", ErrorCategory.NotSpecified, _name));
            }

        }

    }

    [Cmdlet(VerbsCommon.Remove, "CMDBObject", DefaultParameterSetName = "No")]
    public class RemoveCMDBObjectCommand : EntityTypes
    {
        // The adapted EMO

        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public PSObject CardObject { get; set; }

        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            try
            {
                var o = (card)CardObject.BaseObject;
                var returnval = _clientconnection.deleteCard(o.className,o.id);
                WriteObject(returnval);
            }
            catch (Exception e)
            {
                WriteError(new ErrorRecord(e, "Unknown error", ErrorCategory.NotSpecified, ""));
            }

        }
    }

    [Cmdlet(VerbsCommon.Set, "CMDBObject", DefaultParameterSetName = "No")]
    public class SetCMDBObjectCommand : EntityTypes
    {
        // The adapted EMO

        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public PSObject CardObject { get; set; }
        [Parameter(Position = 1, Mandatory = false, ValueFromPipeline = true)]
        public Hashtable PropertyHashtable { get; set; }

        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            try
            {
                var o = (card)CardObject.BaseObject;
                if (o != null)
                {
                    //var resclass2 = _clientconnection.getClassSchema(o.className);
                    var resclass = _clientconnection.getAttributeList(o.className);
                    WriteVerbose("Class: " + o.className);
                    if (resclass != null)
                    {
                        var ht = new Hashtable(StringComparer.OrdinalIgnoreCase);
                        foreach (var prop in resclass)
                        {
                            try
                            {
                                WriteVerbose("Property in Class: " + prop.name);
                                ht.Add(prop.name, prop);
                            }
                            catch (Exception e)
                            {
                                WriteError(new ErrorRecord(e,
                                                           "property '" + prop.name +
                                                           "' has already been added to collection",
                                                           ErrorCategory.InvalidOperation, prop));
                            }
                        }

                        var cardht = new Hashtable(StringComparer.OrdinalIgnoreCase);
                        foreach (attribute d in o.attributeList)
                        {
                            try
                            {
                                cardht.Add(d.name, d);
                            }
                            catch (Exception e)
                            {
                                WriteError(new ErrorRecord(e,
                                                           "property '" + d.name +
                                                           "' has already been added to collection cardht",
                                                           ErrorCategory.InvalidOperation, d));
                            }
                        }

                        var ht2 = new Hashtable(StringComparer.OrdinalIgnoreCase);

                        foreach (var pr in CardObject.Properties)
                        {
                            WriteVerbose(string.Format("Prop Name: {0} PropValue: {1}", pr.Name, pr.Value));
                            ht2.Add(pr.Name, pr.Value);
                        }
                        // now go through the hashtable that has the values we we want to use and
                        // assigned them into the new values
                        foreach (string s in ht2.Keys)
                        {
                            if (ht.ContainsKey(s))
                            {
                                var p = ht[s] as attributeSchema;
                                if (p != null)
                                {
                                    WriteVerbose("Assign Value");
                                    AssignNewValue(this, p, o, ht2[s]);

                                }
                            }
                        }
                        WriteVerbose("Update Card now");
                        var returnval = _clientconnection.updateCard(o);
                        WriteObject(returnval);
                    }
                }


            }
            catch (Exception e)
            {
                WriteError(new ErrorRecord(e, "Unknown error", ErrorCategory.NotSpecified, ""));
            }

        }
    }

    [Cmdlet(VerbsCommon.New, "CMDBObject", DefaultParameterSetName = "No")]
    public class NewCMDBObjectCommand : EntityTypes
    {
        // Parameters
        private string _name = "Class";
        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public string ClassName
        {
            get { return _name; }
            set { _name = value; }
        }
        [Parameter(Position = 1, Mandatory = true, ValueFromPipeline = true)]
        public Hashtable PropertyHashtable { get; set; }

        [Parameter]
        public SwitchParameter PassThru { get; set; }

        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            try
            {
                var resclass = _clientconnection.getClassSchema(_name);
                if (resclass != null)
                {
                    if (!resclass.superClass)
                    {
                        var o = new card {className = _name};

                        var ht = new Hashtable(StringComparer.OrdinalIgnoreCase);
                        var propclass = _clientconnection.getAttributeList(_name);
                        foreach (var prop in propclass)
                        {
                            try
                            {
                                WriteVerbose("Property in Class: " + prop.name);
                                ht.Add(prop.name, prop);
                            }
                            catch (Exception e)
                            {
                                WriteError(new ErrorRecord(e,
                                                           "property '" + prop.name +
                                                           "' has already been added to collection",
                                                           ErrorCategory.InvalidOperation, prop));
                            }
                        }

                        // now go through the hashtable that has the values we we want to use and
                        // assigned them into the new values
                        foreach (string s in PropertyHashtable.Keys)
                        {
                            if (!ht.ContainsKey(s))
                            {
                                WriteError(new ErrorRecord(new SystemException(s), "property not found on object",
                                                           ErrorCategory.NotSpecified, o));
                            }
                            else
                            {

                                var p = ht[s] as attributeSchema;
                                if (p != null)
                                {
                                    WriteVerbose(p.name);
                                    AssignNewValue(this, p, o, PropertyHashtable[s]);
                                }
                            }
                        }
                        //Create Card
                        var rcard = _clientconnection.createCard(o);
                        if (PassThru)
                        {
                            var r = _clientconnection.getCard(_name, rcard, null);
                            WriteObject(AdaptCardObject(this, r));
                        }
                    }
                    else
                    {
                        WriteError(new ErrorRecord(new SystemException("Cannot create Object for SuperClass"), "Unknown error", ErrorCategory.NotSpecified, _name));
                    }
                }
            }
            catch (Exception e)
            {

                WriteError(new ErrorRecord(e, "Unknown error", ErrorCategory.NotSpecified, _name));
            }

        }

    }

    [Cmdlet(VerbsCommon.Get, "CMDBReference", DefaultParameterSetName = "No")]
    public class GetCMDBReferenceCommand : EntityTypes
    {
        // Parameters
        private string _name = "Class";
        [Parameter(Position = 0, Mandatory = false, ValueFromPipeline = true)]
        public string ClassName
        {
            get { return _name; }
            set { _name = value; }
        }

        [Parameter(Position = 1, Mandatory = false, ValueFromPipeline = true)]
        public string Filter { get; set; }

        [Parameter(Position = 2, Mandatory = false, ValueFromPipeline = true)]
        public string FullText { get; set; }

        public GetCMDBReferenceCommand()
        {
            Filter = null;
        }


        protected override void ProcessRecord()
        {

            base.BeginProcessing();

            reference[] refs = null;
            if (Filter != null && Filter.Length > 1)
            {
                WriteVerbose("Query with Filter");
                var query = ConvertFilterToQuery(Filter);
                if (query != null)
                {
                    refs = _clientconnection.getReference(_name, query, null, 0, 0, null, null);
                }
            }
            else if (FullText != null)
            {
                refs = _clientconnection.getReference(_name, null, null, 0, 0, FullText, null);
            }
            else
            {
                WriteVerbose("Query Objects");
                refs = _clientconnection.getReference(_name, null, null, 0, 0, null, null);
            }

            if (refs != null)
            {
                foreach (var c in refs)
                {
                    WriteObject(c);
                }
            }






        }
    }

    [Cmdlet(VerbsCommon.Get, "CMDBObject", DefaultParameterSetName = "No")]
    public class GetCMDBObjectCommand : EntityTypes
    {
        // Parameters
        private string _name = "Class";
        [Parameter(Position = 0, Mandatory = false, ValueFromPipeline = true)]
        public string ClassName
        {
            get { return _name; }
            set { _name = value; }
        }
        [Parameter(Position = 1, Mandatory = false, ValueFromPipeline = true)]
        public int Id { get; set; }

        [Parameter(Position = 2, Mandatory = false, ValueFromPipeline = true)]
        public string Filter { get; set; }

        [Parameter(Position = 3, Mandatory = false, ValueFromPipeline = true)]
        public string FullText { get; set; }

        public GetCMDBObjectCommand()
        {
            Filter = null;
            Id = 0;
        }


        protected override void ProcessRecord()
        {

            base.BeginProcessing();


            if (Id != 0)
                {
                    var card1 = _clientconnection.getCard(_name, Id, null);
                    WriteObject(AdaptCardObject(this, card1));
                }
                else
                {
                    
                    cardList cards = null;
                    if (Filter != null && Filter.Length > 1  )
                    {
                        WriteVerbose("Query with Filter");
                        var query = ConvertFilterToQuery(Filter);
                        if (query != null)
                        {
                            cards = _clientconnection.getCardList(_name, null, query, null, 0, 0, null, null);
                        }
                    }
                    else if (FullText != null)
                    {
                        cards = _clientconnection.getCardList(_name, null, null, null, 0, 0, FullText, null);
                    }
                    else
                    {
                        WriteVerbose("Query Objects");
                        cards = _clientconnection.getCardList(_name, null, null, null, 0, 0, null, null);
                    }

                    if (cards != null && cards.cards != null)
                    {
                        foreach (var c in cards.cards)
                        {
                            var attrarray = _clientconnection.getAttributeList(c.className).ToList();
                            foreach (attribute n1 in c.attributeList)
                            {
                                foreach (attributeSchema as1 in attrarray){
                                    if (as1.type == "REFERENCE" && as1.name == n1.name)
                                    {
                                        n1.value = n1.code;
                                        break;
                                    }
                                }
                            }
                            WriteObject(AdaptCardObject(this,c));
                        }
                    }

                }


 
        }

    }

    [Cmdlet(VerbsCommon.Get, "CMDBObjectHistory", DefaultParameterSetName = "No")]
    public class GetCMDBObjectHistoryCommand : EntityTypes
    {
        // Parameters
        private string _name = "Class";
        [Parameter(Position = 0, Mandatory = false, ValueFromPipeline = true)]
        public string ClassName
        {
            get { return _name; }
            set { _name = value; }
        }
        [Parameter(Position = 1, Mandatory = false, ValueFromPipeline = true)]
        public int Id { get; set; }


        public GetCMDBObjectHistoryCommand()
        {
            Id = 0;
        }


        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            var cardhistory1 = _clientconnection.getCardHistory(_name, Id, 0, 0);
            WriteObject(cardhistory1);

        }

    }

    [Cmdlet(VerbsCommon.Remove, "CMDBLookupBy", DefaultParameterSetName = "Id")]
    public class RemoveCMDBLookupByCommand : EntityTypes
    {
        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Id")]
        public int Id { get; set; }
        // Parameters
        private string _name = "";
        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Code")]
        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }
        private string _code = "";
        [Parameter(Position = 1, Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Code")]
        public string Code
        {
            get { return _code; }
            set { _code = value; }
        }

        [Parameter(Position = 2, Mandatory = false, ParameterSetName = "Code")]
        public SwitchParameter PassThru { get; set; }


        public RemoveCMDBLookupByCommand()
        {
            Id = 0;

        }
        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            try
            {
                if (Id != 0)
                {

                    var lookup = _clientconnection.deleteLookup(Id);
                    WriteObject(lookup);
                }
                if (!string.IsNullOrEmpty(Code) && !string.IsNullOrEmpty(Name))
                {
                    var lookups = _clientconnection.getLookupListByCode(Name, Code, PassThru);
                    foreach (var lookup in lookups)
                    {
                        WriteVerbose("Remove Lookup: " + lookup.id);
                        var l = _clientconnection.deleteLookup(lookup.id);
                        WriteObject(l);
                    }
                }
            }
            catch (Exception e)
            {

                WriteError(new ErrorRecord(e, "Unknown error", ErrorCategory.NotSpecified, _name));
            }

        }

    }

    [Cmdlet(VerbsCommon.Get, "CMDBLookupBy",DefaultParameterSetName = "Id")]
    public class GetCMDBLookupByCommand : EntityTypes
    {
        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true,ParameterSetName = "Id")]
        public int Id { get; set; }
        // Parameters
        private string _name = "";
        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Code")]
        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }
        private string _code = "";
        [Parameter(Position = 1, Mandatory = true, ValueFromPipeline = true,ParameterSetName = "Code")]
        public string Code
        {
            get { return _code; }
            set { _code = value; }
        }

        [Parameter(Position = 2,Mandatory = false,ParameterSetName = "Code")]
        public SwitchParameter PassThru { get; set; }

 
        public GetCMDBLookupByCommand()
        {
            Id = 0;

        }
        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            try
            {
                if (Id != 0)
                {
         
                    var lookups = _clientconnection.getLookupById(Id);
                    WriteObject(lookups);
                }
                if(!string.IsNullOrEmpty(Code) && !string.IsNullOrEmpty(Name)  )
                {
                    var lookups = _clientconnection.getLookupListByCode(Name, Code, PassThru);
                    WriteObject(lookups);
                }
            }
            catch (Exception e)
            {

                WriteError(new ErrorRecord(e, "Unknown error", ErrorCategory.NotSpecified, _name));
            }

        }

    }

    [Cmdlet(VerbsCommon.Get, "CMDBLookupList", DefaultParameterSetName = "No")]
    public class GetCMDBLookupListCommand : EntityTypes
    {
        // Parameters
        private string _name = "";
        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true)]
        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }
        protected override void ProcessRecord()
        {
            base.BeginProcessing();
            try
            {
                var lookups = _clientconnection.getLookupList(_name, null, false);
                WriteObject(lookups);
            }
            catch (Exception e)
            {

                WriteError(new ErrorRecord(e, "Unknown error", ErrorCategory.NotSpecified, _name));
            }

        }

    }

    #region Sessions cmdlets
    [Cmdlet("New","CMDBSession",DefaultParameterSetName="URLANDCRED")]
    public class NewCMDBSession : CMDBCmdletBase
    {
        [Parameter]
        public SwitchParameter PassThru { get; set; }
        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            if (PassThru) { WriteObject(_clientconnection); }
        }
    }

    [Cmdlet("Get", "CMDBSession")]
    public class GetCMDBSession : PSCmdlet
    {
        private string _uri = "";
        [Parameter(Position = 0, ValueFromPipeline = true, Mandatory = true)]
        [ValidateNotNullOrEmpty]
        public string URL
        {
            get { return _uri; }
            set { _uri = value; }
        }

        private List<string> l = null;
        protected override void ProcessRecord()
        {
            l = ConnectionHelper.GetClientConnectionsList(URL);
            if (l != null)
            {
                foreach (string n in l)
                {
                    WriteObject(ConnectionHelper.GetClientConnection(n));
                }
            }
        }
    }

    [Cmdlet("Remove", "CMDBSession")]
    public class RemoveCMDBSession : PSCmdlet
    {
        private PrivateClient _client;
        [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
        public PrivateClient Client
        {
            get { return _client; }
            set { _client = value; }
        }
        protected override void ProcessRecord()
        {
            ConnectionHelper.RemoveClientConnection(Client.Endpoint.ListenUri.OriginalString);
        }
    }

    public sealed class ConnectionHelper
    {
        private static Hashtable ht;

        public static void SetClientConnection(PrivateClient client)
        {
            if (ht == null) { ht = new Hashtable(StringComparer.OrdinalIgnoreCase); }
            if (!ht.ContainsKey(client.Endpoint.Address.Uri.ToString()))
            {
                ht.Add(client.Endpoint.Address.Uri.ToString(), client);
            }
        }
        public static List<string> GetClientConnectionsList(string re)
        {
            var r = new Regex(re, RegexOptions.IgnoreCase);

            return (from string k in ht.Keys where r.Match(k).Success select k).ToList();

        }
        public static PrivateClient GetClientConnection(string uri)
        {
            if (ht == null) { ht = new Hashtable(StringComparer.OrdinalIgnoreCase); }
            if (ht.ContainsKey(uri))
            {
                return (PrivateClient)ht[uri];
            }
            return null;
        }
        public static void RemoveClientConnection(string uri)
        {
            if (ht == null) { ht = new Hashtable(StringComparer.OrdinalIgnoreCase); }
            if (ht.ContainsKey(uri))
            {
                ht.Remove(uri);
            }
        }




        private static PrivateClient NewClient(string uri, PSCredential credential)
        {
            CustomBinding custBinding = new CustomBinding();
            MtomMessageEncodingBindingElement elmtom = new MtomMessageEncodingBindingElement();
            elmtom.MaxWritePoolSize = 2147483647;
            elmtom.MaxBufferSize = 2147483647;
            elmtom.MessageVersion = MessageVersion.Soap12;


            var ssbe = SecurityBindingElement.CreateUserNameOverTransportBindingElement();
            ssbe.AllowInsecureTransport = true;
            ssbe.IncludeTimestamp = false;
            ssbe.MessageSecurityVersion = MessageSecurityVersion.WSSecurity10WSTrustFebruary2005WSSecureConversationFebruary2005WSSecurityPolicy11BasicSecurityProfile10;
            var htt = new HttpTransportBindingElement();

            htt.MaxBufferPoolSize = 2147483647;
            htt.MaxReceivedMessageSize = 2147483647;
            htt.MaxBufferSize = 2147483647;
            EndpointAddress n = new EndpointAddress(uri);

            custBinding.Elements.Add(ssbe);
            custBinding.Elements.Add(elmtom);
            custBinding.Elements.Add(htt);

            var client = new PrivateClient(custBinding, n);
                if (credential != null)
                {
                    
                    if (client.ClientCredentials != null)
                    {
                        
                        client.ClientCredentials.UserName.UserName = credential.GetNetworkCredential().UserName;
                        client.ClientCredentials.UserName.Password = SecureStringToString(credential.Password);
                    }
                }
                else
                {
                    if (client.ClientCredentials != null)
                    {
                       // client.ClientCredentials.UserName.UserName = "admin";
                        //client.ClientCredentials.UserName.Password = "admin";
                    }
                }
            return client;
        }
     
        public static PrivateClient GetClientConnection(string uri, PSCredential credential)
        {
            if (ht == null) { ht = new Hashtable(StringComparer.OrdinalIgnoreCase); }
            if (!ht.ContainsKey(uri))
            {
                var client = NewClient(uri, credential);
                ht.Add(uri, client);
            }
            if (((PrivateClient)ht[uri]).State == CommunicationState.Faulted && credential != null)
            {
                var client = NewClient(uri, credential);
                ht[uri] =  client;
            }
            return ht[uri] as PrivateClient;
            
            
        }

        private static String SecureStringToString(SecureString value)
        {
            IntPtr bstr = Marshal.SecureStringToBSTR(value);

            try
            {
                return Marshal.PtrToStringBSTR(bstr);
            }
            finally
            {
                Marshal.FreeBSTR(bstr);
            }
        }

    }
    #endregion

}
