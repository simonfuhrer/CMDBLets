using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Management.Automation;
using cmdblets.CMDBUILD;
namespace cmdblets
{
    public class CMDBCmdletBase : PSCmdlet
    {
        private string _uri = "";
        internal PrivateClient _clientconnection;
        [Parameter(Mandatory = false, ParameterSetName = "URL", HelpMessage = "The URL to use for the connection to the CMDBUILD Webservice. http://192.168.2.224:8080/cmdb")]
        [Parameter(Mandatory = true, ParameterSetName = "URLANDCRED", HelpMessage = "The URL to use for the connection to the CMDBUILD Webservice. http://192.168.2.224:8080/cmdb")]
        [ValidateNotNullOrEmpty]
        public string URL
        {
            get { return _uri; }
            set { _uri = value; }
        }

        private PSCredential _credential = null;
        [Parameter(Mandatory = true, ParameterSetName = "URLANDCRED", HelpMessage = "Credentials used to connect to")]
        [ValidateNotNullOrEmpty]
        public PSCredential Credential
        {
            get { return _credential; }
            set { _credential = value; }
        }
        private PrivateClient _cmdbSession = null;
        [Parameter(Mandatory = true, ParameterSetName = "SESSION", HelpMessage = "A connection to a CMDBUILD Webservice")]
        public PrivateClient CMDBSession
        {
            get { return _cmdbSession; }
            set { _cmdbSession = value; }
        }

        protected override void BeginProcessing()
        {
            try
            {

                if (CMDBSession != null)
                {
                    ConnectionHelper.SetClientConnection(CMDBSession);
                    _clientconnection = CMDBSession;
                }
                else
                {
                    if (string.IsNullOrEmpty(URL))
                    {
                        var DefaultCMDBURL = SessionState.PSVariable.Get("DefaultCMDBURL");
                        if (DefaultCMDBURL == null)
                        {
                            ThrowTerminatingError(
                                new ErrorRecord(
                                    new Exception("No Session Object found! Please initaite first a Session Object!"),
                                    "GenericMessage", ErrorCategory.ObjectNotFound, ""));
                        }
                        _uri = DefaultCMDBURL.Value.ToString();
                    }
                    var uri = string.Format("{0}/services/soap/Private", _uri);
                    _clientconnection = ConnectionHelper.GetClientConnection(uri, _credential);
                     SessionState.PSVariable.Set("DefaultCMDBURL",_uri);

                }
            }
            catch (Exception e)
            {

                ThrowTerminatingError(new ErrorRecord(e, "GenericMessage", ErrorCategory.InvalidOperation, _uri)
               );
            }



        }


    }
}
