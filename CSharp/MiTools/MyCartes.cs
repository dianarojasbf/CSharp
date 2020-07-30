using Cartes;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Xml;
using Microsoft.Win32;
using System.Windows.Forms;

//////////////////////
// 2020/06/25 13:20
//    v2.1
//////////////////////

namespace MiTools
{
    public abstract class MyCartes : MyObject // This abstract class is the basis for working and grouping methods for Cartes.
    {
        protected const string EXIT_ABORT = "Abort";
        protected const string EXIT_ERROR = "KO";
        protected const string EXIT_OK = "OK";

        private static Version fVersion = null;
        private Version flVersion = null;
        private NumberFormatInfo fDoubleFormatProvider = null;
        private string fAbort; // Cartes Script Variable for know when the process musts abort

        public MyCartes(string csvAbort) // "Owner" is Cartes. "csvAbort" is a Cartes variable that when valid one will indicate to the instance that it must abort.
        {
            fAbort = ToString(csvAbort).Trim();
            flVersion = null;
        }

        internal void Load()
        {
            MergeLibrariesAndLoadVariables();
        }
        internal void UnLoad()
        {
            UnMergeLibrariesAndUnLoadVariables();
        }

        private Version getNeededRPASuiteVersionP() // It returns the version of RPA Suite needed by this library
        {
            try
            {
                if (flVersion == null)
                {
                    if (!Version.TryParse(ToString(getNeededRPASuiteVersion()), out flVersion))
                        throw new Exception(getNeededRPASuiteVersion() + " is not a valid version number.");
                }
            }catch(Exception e)
            {
                forensic("MyCartes.getNeededRPASuiteVersionP", e);
                throw;
            }
            return flVersion;
        }
        private bool GetIsRPASuiteInstalled()
        {
            bool result;
            try
            {
                Version v = CurrentRPASuiteVersion;
                result = v.ToString().Length > 0;
            }
            catch
            {
                result = false;
            }
            return result;
        }

        protected abstract CartesObj getCartes();
        protected abstract string getCartesPath();  // It returns the file of Cartes
        protected abstract void MergeLibrariesAndLoadVariables();  // Rewrite this method to load the libraries and Cartesa variables that your class handles.
        protected abstract void UnMergeLibrariesAndUnLoadVariables(); // Rewrite this method to unload the libraries and Cartesa variables that your class handles.
        protected virtual Version getCurrentRPASuiteVersion()  // It returns the version of RPA Suite
        {
            if (fVersion == null)
            {
                object lvalue;

                lvalue = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Cartes", "Product Version", null);
                if (lvalue == null)
                    lvalue = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Cartes", "Product Version", null);
                if (!Version.TryParse(ToString(lvalue), out fVersion))
                    throw new Exception("RPA Suite is not installed. Please install RPA Suite.");
            }
            return fVersion;
        }
        protected virtual string getNeededRPASuiteVersion() // It returns a string with the version of RPA Suite needed by this library
        {
            return "3.2.1.0";
        }
        protected virtual void CheckRPASuiteVersion() // It checks if the current version and needed are OK
        {
            if (CurrentRPASuiteVersion.CompareTo(NeededRPASuiteVersion) < 0)
                throw new Exception("You need RPA Suite v" + NeededRPASuiteVersion + " or higher.");
        }
        protected virtual NumberFormatInfo getDoubleFormatProvider()
        {
            if (fDoubleFormatProvider == null)
            {
                fDoubleFormatProvider = new NumberFormatInfo();
                fDoubleFormatProvider.NumberDecimalSeparator = ",";
                fDoubleFormatProvider.NumberGroupSeparator = ".";
                fDoubleFormatProvider.NumberGroupSizes = new int[] { 3 };
            }
            return fDoubleFormatProvider;
        }
        protected abstract string getProjectId(); // Returns the ID of the loaded project in Cartes. If Cartes does not have a loaded project, it returns the empty string.
        protected virtual void reset(RPAComponent component) // Reset the API of te component
        {
            cartes.reset(component.api());
        }
        protected void reset(RPAWin32Component component) // Reset the API of te component
        {
            reset((RPAComponent)component);
        }
        protected void reset(RPAWin32Automation component) // Reset the API of te component
        {
            reset((RPAComponent)component);
        }
        protected void reset(RPAMSHTMLComponent component) // Reset the API of te component
        {
            reset((RPAComponent)component);
        }
        protected bool isVariable(string VariableName)  // If a variable-component exists in the rpa project, returns true
        {
            try
            {
                return Execute("isVariable(\"" + VariableName + "\");") == "1";
            }
            catch
            {
                return false;
            }
        }
        protected virtual RPAComponent GetComponent<RPAComponent>(string variablename) where RPAComponent : class, IRPAComponent
        {
            return cartes.GetComponent<RPAComponent>(variablename);
        }
        protected virtual string Execute(string command) // It Executes a Cartes Script in cartes.execute and check if errors
        {
            string result;

            result = ToString(cartes.Execute(command));
            if ((cartes.LastError() != null) && (cartes.LastError().Length > 0))
                throw new Exception(cartes.LastError());
            else if (result == null) return "";
            else return result;
        }
        protected virtual new void forensic(string message) // It writes "message" in the swarm log and in the windows event viewer.
        {
            cartes.forensic(message);
        }
        protected virtual new void forensic(string message, Exception e)
        {
            forensic(message + "\r\n" + e.Message);
        }
        protected virtual void CheckAbort() // It checks if the variable to abort is 1 to throw an exception
        {
            if ((Abort.Length > 0) && (cartes.Execute(Abort + ".value;") == "1"))
                throw new MyException(EXIT_ABORT, "Abort by user.");
        }
        protected virtual void AdjustWindow(RPAWin32Component component, int width, int height) // Adjusts the main component window to the indicated size.
        {
            RPAWin32Component lpWindow = (RPAWin32Component)component.getComponentRoot();

            if (lpWindow.componentexist(0) == 1)
            {
                if (StringIn(lpWindow.WindowState, "Minimized", "Maximized") || (lpWindow.Visible == 0))
                {
                    lpWindow.Show("Restore");
                }
                if ((lpWindow.width != width) || (lpWindow.height != height))
                    lpWindow.ReSize(width, height);
                if ((lpWindow.x != 0) || (lpWindow.y != 0))
                    lpWindow.Move(0, 0);
            }
        }
        protected virtual void scrollUp(int mouseX, int mouseY, RPAWin32Accessibility component)
        {
            RPAParameters parametros = new RPAParameters();
            DateTime timeout;
            int n, y;

            try
            {
                parametros.item[0] = mouseX.ToString();
                parametros.item[1] = mouseY.ToString();
                component.focus();
                component.doroot("SetMouse", parametros);
                n = 0;
                y = component.y;
                timeout = DateTime.Now.AddSeconds(30);
                while (component.OffScreen == 1)
                {
                    if (timeout < DateTime.Now) throw new Exception("Timeout waiting by scrolling.");
                    CheckAbort();
                    cartes.balloon("Scroll...");
                    component.doroot("SetMouse", parametros);
                    component.Up();
                    Thread.Sleep(500);
                    if (n > 3) throw new Exception("I can not scroll up.");
                    if (y == component.y) n++;
                    else n = 0;
                }
            }
            catch (Exception e)
            {
                cartes.forensic("MyCartes.scrollUp\r\n" + e.Message);
                throw;
            }
        }
        protected virtual void scrollDown(int mouseX, int mouseY, RPAWin32Accessibility component)
        {
            RPAParameters parameters = new RPAParameters();
            DateTime timeout;
            int n, y;

            try
            {
                parameters.item[0] = mouseX.ToString();
                parameters.item[1] = mouseY.ToString();
                component.focus();
                component.doroot("SetMouse", parameters);
                n = 0;
                y = component.y;
                timeout = DateTime.Now.AddSeconds(30);
                while (component.OffScreen == 1)
                {
                    if (timeout < DateTime.Now) throw new Exception("Timeout waiting by scrolling.");
                    CheckAbort();
                    cartes.balloon("Scroll...");
                    component.doroot("SetMouse", parameters);
                    component.down();
                    Thread.Sleep(500);
                    if (n > 3) throw new Exception("I can not scroll down.");
                    if (y == component.y) n++;
                    else n = 0;
                }
            }
            catch (Exception e)
            {
                cartes.forensic("MyCartes.scrollDown(RPAWin32Accessibility)\r\n" + e.Message);
                throw;
            }
        }
        protected virtual void scrollDown(int mouseX, int y, int height, RPAWin32Component component)
        {
            RPAParameters parametros = new RPAParameters();
            DateTime timeout;

            try
            {
                parametros.item[0] = mouseX.ToString();
                parametros.item[1] = (y + 2).ToString();
                component.focus();
                component.doroot("SetMouse", parametros);
                timeout = DateTime.Now.AddSeconds(20);
                while (component.y < y)
                {
                    if (timeout < DateTime.Now) throw new Exception("Timeout waiting by scrolling.");
                    CheckAbort();
                    cartes.balloon("Scroll up...");
                    component.Up();
                    Thread.Sleep(500);
                }
                while (y + height < component.y + component.height)
                {
                    if (timeout < DateTime.Now) throw new Exception("Timeout waiting by scrolling.");
                    CheckAbort();
                    cartes.balloon("Scroll down...");
                    component.down();
                    Thread.Sleep(500);
                }
            }
            catch (Exception e)
            {
                cartes.forensic("MyCartes.scrollDown(RPAWin32Component)\r\n" + e.Message);
                throw;
            }
        }
        protected virtual void scrollDown(int mouseX, int y, int height, RPAWin32Accessibility component)
        {
            scrollDown(mouseX, y, height, (RPAWin32Component)component);
        }
        protected virtual bool ComponentsExist(int seconds, params RPAComponent[] components) /* The method waits for the indicated
              seconds until one of the components exists. If any of the components exist, returns true. */
        {
            DateTime timeout;
            bool exit, result = false;
            string lsApi = "";

            try
            {
                exit = false;
                timeout = DateTime.Now.AddSeconds(seconds);
                do
                {
                    CheckAbort();
                    foreach(RPAComponent component in components)
                    {
                            if (component.componentexist(0) == 1)
                            {
                                result = true;
                                exit = true;
                                break;
                            }
                            else if (lsApi.Length == 0)
                                lsApi = component.api();
                    }
                    if (timeout < DateTime.Now) exit = true;
                    else if (!exit)
                    {
                        Thread.Sleep(400);
                        if (lsApi.Length > 0)
                            cartes.reset(lsApi);
                    }
                } while (!exit);
            }
            catch (Exception e)
            {
                forensic("MyCartes::ComponentsExist", e);
                throw;
            }
            return result;
        }
        protected bool ComponentsExist(int seconds, params RPAWin32Component[] components)
        {
            List<RPAComponent> parametros = new List<RPAComponent>();

            foreach (RPAWin32Component win in components)
            {
                parametros.Add((RPAComponent)win);
            }
            return ComponentsExist(seconds, parametros.ToArray());
        }
        protected virtual string WaitForCartesMethodValue(RPAComponent component, string method, int seconds) /* It waits until the indicated method has
            value and returns it. Once the waiting time has been exceeded, it throws an exception. */
        {
            RPAParameters parameters = new RPAParameters();
            DateTime timeout;
            string result;

            try
            {
                timeout = DateTime.Now.AddSeconds(seconds);
                do
                {
                    CheckAbort();
                    cartes.reset(component.api());
                    Thread.Sleep(400);
                    if (timeout < DateTime.Now) throw new Exception("Timeout");
                    else result = ToString(component.Execute(method, parameters));
                } while (result.Length == 0);
                return result;
            }
            catch (Exception e)
            {
                forensic("MyCartes::WaitForCartesMethodValue", e);
                throw;
            }
        }
        protected virtual string WaitForCartesMethodValue(RPAWin32Accessibility component, string method, int seconds)
        {
            return WaitForCartesMethodValue((RPAComponent)component, method, seconds);
        }
        protected virtual void AssignValueInsistently(DateTime timeout, RPAWin32Component component, string value, bool typed = false) /* Assign the indicated value to the
            component. If it does not succeed, it will insist until the system time exceeds "timeot". */
        {
            do
            {
                if (typed)
                    try
                    {
                        component.TypeFromClipboardCheck(value, 0, 0);
                    }
                    catch 
                    {
                        component.TypeWordCheck(value, 0, 0);
                    }
                else component.Value = value;
                CheckAbort();
                Thread.Sleep(1000);
                reset(component);
                if (ToString(component.Value).ToLower() == value.ToLower()) break;
                else
                {
                    if (timeout < DateTime.Now) throw new Exception("I can't assign the value \"" + value + "\" to the component.");
                    Thread.Sleep(1000);
                }
            } while (true);
        }
        protected virtual double SimilarStrings(string a, string b)
        {
            NumberFormatInfo fDoubleFormat = new NumberFormatInfo();
            string sresult;
            double dresult;

            try
            {
                fDoubleFormat.NumberDecimalSeparator = ".";
                fDoubleFormat.NumberGroupSeparator = ",";
                fDoubleFormat.NumberGroupSizes = new int[] { 3 };
                sresult = cartes.Execute("similarstrings(\"spa\", \"" + a.Replace("\"", "\"\"") + "\", \"" + b.Replace("\"", "\"\"") + "\");");
                dresult = Convert.ToDouble(sresult, fDoubleFormat);
            }
            catch(Exception e)
            {
                forensic("MyCartes.SimilarStrings", e);
                throw;
            }
            return dresult;
        }

        public override double ToDouble(string value)
        {
            return Convert.ToDouble(value, getDoubleFormatProvider());
        }

        public bool IsRPASuiteInstalled
        {
            get { return GetIsRPASuiteInstalled(); }
        }  // Read Only. It returns if RPA Suite is installed
        public Version CurrentRPASuiteVersion
        {
            get { return getCurrentRPASuiteVersion(); }
        }  // Read Only. It returns the version of RPA Suite
        public Version NeededRPASuiteVersion
        {
            get { return getNeededRPASuiteVersionP(); }
        }  // Read Only. It returns the version of RPA Suite needed by this library
        public string CartesPath
        {
            get
            {
                return getCartesPath();
            }
        } // Read. It returns the file of Cartes
        public CartesObj cartes
        {
            get { return getCartes(); }
        }  // Read Only
        public string ProjectId
        {
            get { return getProjectId(); }
        }  // Read Only. Returns the ID of the loaded project in Cartes. If Cartes does not have a loaded project, it returns the empty string.
        public string Abort
        {
            get { return fAbort; }
        }  // Read Only
        public NumberFormatInfo DoubleFormatProvider
        {
            get { return getDoubleFormatProvider(); }
        }  // Read Only
    }

    public abstract class MyCartesProcess : MyCartes // This abstract class allows you to create processes using APIS from MyCartesAPI.
    {
        protected const string EXIT_SETTINGS_KO = "Settings_" + EXIT_ERROR;
        private static string fCartesPath = null;
        private CartesObj fCartes;
        private List<MyCartes> apis;
        private string fFileSettings;
        private bool fShowAbort, fVisibleMode;
        private RPADataString frpaAbort = null;
        protected SmtpClient fSMTP;

        public MyCartesProcess(string csvAbortar) : base(csvAbortar)
        {
            fCartes = null;
            apis = new List<MyCartes>();
            fFileSettings = null;
            fShowAbort = true;
            fVisibleMode = true;
            fSMTP = null;
        }
  
        internal void AddAPI(MyCartes api)
        {
            if (apis.IndexOf(api) < 0)
            {
                apis.Add(api);
                if (ProjectId.Length > 0)
                    api.Load();
            }
        }
        internal void DeleteAPI(MyCartes api)
        {
            apis.Remove(api);
        }
 
        private void LoadConfiguration() // It loads the process configuration.   
        {
            XmlDocument lpXmlCfg = null;

            try
            {
                lpXmlCfg = new XmlDocument();
                lpXmlCfg.Load(SettingsFile);
                LoadConfiguration(lpXmlCfg);
            }catch(Exception e)
            {
                if (e is MyException m) throw m;
                else throw new MyException(EXIT_SETTINGS_KO, e.Message);
            }
        }
 
        protected override CartesObj getCartes()
        {
            if (fCartes == null)
            {
                if ((CartesPath.Length > 0) && File.Exists(CartesPath))
                {
                    bool ok;
                    string CartesName = Path.GetFileNameWithoutExtension(CartesPath);
                    System.Diagnostics.Process current = System.Diagnostics.Process.GetCurrentProcess();
                    System.Diagnostics.Process[] ap = System.Diagnostics.Process.GetProcessesByName(CartesName);
                    ok = false;
                    foreach (System.Diagnostics.Process item in ap)
                    {
                        if (item.SessionId == current.SessionId)
                        {
                            ok = true;
                            break;
                        }
                    }
                    if (!ok)
                    {
                        System.Diagnostics.Process.Start(CartesPath);
                        System.Threading.Thread.Sleep(3000);
                    }
                    fCartes = new CartesObj();
                }
                else throw new Exception("Cartes is not installed. Please install Robot Cartes from the RPA Suite installation.");
            }
            return fCartes;
        }
        protected override string getCartesPath()  
        {
            object InstallPath;

            if (fCartesPath == null)
            {
                InstallPath = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Cartes", "Cartes Client", null);
                if (InstallPath == null)
                    InstallPath = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Cartes", "Cartes Client", null);
                fCartesPath = ToString(InstallPath);
            }
            return fCartesPath;
        }
        protected override void MergeLibrariesAndLoadVariables()
        {
            try
            {
                foreach (MyCartes item in apis)
                    item.Load();
                if (Abort.Length > 0)
                {
                    if (isVariable(Abort)) frpaAbort = (RPADataString)cartes.component(Abort);
                    else throw new MyException(EXIT_ERROR, Abort + " does not exist!");
                }
                else frpaAbort = null;
            }
            catch (Exception e)
            {
                forensic("MyCartesProcess.MergeLibrariesAndLoadVariables", e);
                throw new MyException(EXIT_ERROR, e.Message);
            }
        }
        protected override void UnMergeLibrariesAndUnLoadVariables()
        {
            try
            {
                foreach (MyCartes item in apis)
                    item.UnLoad();
            }
            catch (Exception e)
            {
                forensic("MyCartesProcess.UnMergeLibrariesAndUnLoadVariables", e);
                throw new MyException(EXIT_ERROR, e.Message);
            }
        }
        protected override string getProjectId() 
        {
            return ToString(Execute("ProjectId;"));
        }
        protected SmtpClient getSMTP() // It returns a SMTP Client to send emails.
        {
            if (fSMTP == null)
            {
                fSMTP = new SmtpClient();
                fSMTP.Credentials = new System.Net.NetworkCredential("myaccount@gmail.com", "mypassword");
                fSMTP.Port = 587;
                fSMTP.Host = "smtp.gmail.com";
                fSMTP.EnableSsl = true;
            }
            return fSMTP;
        }
        protected abstract string getRPAMainFile(); // Here you must return the main ".rpa" file
        protected abstract void LoadConfiguration(XmlDocument XmlCfg); // Here you must load the configuration of the process
        protected virtual bool DoInit() // If DoExecute must be invoked, this method returns true.
        {
            return true;
        }
        protected virtual void ShowAbortDialog(RPADataString abort)
        {
            abort.ShowAbortDialog("Press button to abort", "Closing...", "Abort");
        }
        protected abstract void DoExecute(ref DateTime start); // Here you must execute the process. The process have already loaded the configuration
        protected virtual void DoEnd() // This method is invoked after running DoExecute.
        {

        }

        public bool Execute()  // Execute the process. if succesfull return True, else return false
        {
            DateTime start;
            bool result = false;
            string lsMainFile;

            try
            {
                start = DateTime.Now;
                CheckRPASuiteVersion();
                lsMainFile = RPAMainFile;
                if (File.Exists(CurrentPath + "\\" + lsMainFile)) cartes.open(CurrentPath + "\\" + lsMainFile);
                else if (File.Exists(CurrentPath + "\\Cartes\\" + lsMainFile)) cartes.open(CurrentPath + "\\Cartes\\" + lsMainFile);
                else cartes.open(RPAMainFile);
                try
                {
                    cartes.balloon(Execute("Name;"));
                    try
                    {
                        MergeLibrariesAndLoadVariables();
                        LoadConfiguration();
                        try
                        {
                            if (VisibleMode)
                                Execute("visualmode(1);");
                            try
                            {
                                if (DoInit())
                                {
                                    if (ShowAbort && (frpaAbort != null))
                                        ShowAbortDialog(frpaAbort);
                                    DoExecute(ref start);
                                }
                            }
                            finally
                            {
                                DoEnd();
                            }
                        }
                        finally
                        {
                            if (VisibleMode)
                                Execute("visualmode(0);");
                        }
                    }
                    catch (Exception e)
                    {
                        MyException mye;

                        cartes.balloon(e.Message);
                        cartes.forensic(e.Message);
                        if (e is MyException m) mye = m;
                        else mye = new MyException(EXIT_ERROR, e.Message);
                        if (mye.code == EXIT_SETTINGS_KO)
                            cartes.RegisterIteration(start, EXIT_SETTINGS_KO, "<task>\r\n" +
                                                            "  <error>I can not load the configuration file \"" + SettingsFile + "\"</error>\r\n" +
                                                            "  <message>" + e.Message + "</message>\r\n" +
                                                            "</task>", 1);
                        else
                            cartes.RegisterIteration(start, mye.code, "<task>\r\n" +
                                                            "  <data>" + e.Message + "</data>\r\n" +
                                                            "</task>", 1);
                    }
                }
                finally
                {
                    try
                    {
                        UnMergeLibrariesAndUnLoadVariables();
                    }
                    finally
                    {
                        cartes.close();
                    }
                }
                result = true;
            }
            catch (Exception e)
            {
                cartes.balloon(e.Message);
                forensic(e.Message);
#if DEBUG
                MessageBox.Show(e.Message);
#endif
            }
            return result;
        }

        public string SettingsFile
        {
            get
            {
                if (fFileSettings == null)
                    fFileSettings = CurrentPath + "\\settings.xml";
                return fFileSettings;
            }
            set { fFileSettings = value; }
        } // Read & Write
        public bool ShowAbort
        {
            get { return fShowAbort; }
            set { fShowAbort = value; }
        } // Read & Write. It controls the appearance of the window to abort. 
        public bool VisibleMode
        {
            get { return fVisibleMode; }
            set { fVisibleMode = value; }
        } // Read & Write. It controls the visible mode of Carte. 
        public string RPAMainFile
        {
            get { return getRPAMainFile(); }
        } // Read
        public SmtpClient SMTP
        {
            get { return getSMTP(); }
        } // Read
    }

    public abstract class MyCartesAPI : MyCartes // This abstract class allows you to inherit to create application APIs (Chrome, SAP ...) using Cartes for MyCartesProcess.
    {
        private MyCartesProcess fowner;
        private bool fChecked;

        public MyCartesAPI(MyCartesProcess owner) : base(owner.Abort)
        {
            fowner = owner;
            fChecked = false;
            fowner.AddAPI(this);
        }
        ~MyCartesAPI()
        {
            fowner.DeleteAPI(this);
        }
        protected override CartesObj getCartes()
        {
            if (!fChecked)
            {
                CheckRPASuiteVersion();
                fChecked = true;
            }
            return Owner.cartes;
        }
        protected override string getCartesPath()
        {
            return Owner.CartesPath;
        }
        protected override NumberFormatInfo getDoubleFormatProvider()
        {
            return Owner.DoubleFormatProvider;
        }
        protected override string getProjectId()
        {
            return Owner.ProjectId;
        }

        public override double ToDouble(string value)
        {
            return Owner.ToDouble(value);
        }

        public MyCartesProcess Owner
        {
            get { return fowner; }
        }
    }

    public static class CartesObjExtensions
    {
        public static T GetComponent<T>(this CartesObj cartesObj, string variablename) where T : class, IRPAComponent
        {
            IRPAComponent component = cartesObj.component(variablename);

            if (component == null) return null;
            else if (component is T result) return result;
            else throw new Exception(variablename + " is a " + component.ActiveXClass());
        }
        public static RPAWin32Component Root(this RPAWin32Component component)
        {
            return component.getComponentRoot() as RPAWin32Component;
        }
    }
}
