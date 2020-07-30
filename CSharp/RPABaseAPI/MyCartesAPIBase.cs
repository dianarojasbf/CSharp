using Cartes;
using MiTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RPABaseAPI
{
    public abstract class MyCartesAPIBase : MyCartesAPI
    {
        public MyCartesAPIBase(MyCartesProcess owner) : base(owner)
        {
        }
        protected override void UnMergeLibrariesAndUnLoadVariables() // Normally, you don't need to do anything to download the library.
        {
        }
    }

    public class ClassVisualStudio : MyCartesAPIBase
    {
        private static bool loaded = false;
        private RPAWin32Component vsWindow = null;

        public ClassVisualStudio(MyCartesProcess owner) : base(owner)
        {
            MergeLibrariesAndLoadVariables();
        }
        
        protected override void MergeLibrariesAndLoadVariables()
        {
            if (!loaded || !isVariable("$VisualStudio"))
            {
                loaded = cartes.merge(CurrentPath + "\\Cartes\\VisualStudio.rpa") == 1;
            }
            if (vsWindow == null)
            {
                vsWindow = GetComponent<RPAWin32Component>("$VisualStudio");
            }
        }
        protected override void UnMergeLibrariesAndUnLoadVariables()
        {
            loaded = false;
            vsWindow = null;
        }
        public virtual bool Minimice() // Minimice the Visual Studio
        {
            bool result = false;
            RPAWin32Component lpWindow = (RPAWin32Component)vsWindow.getComponentRoot();

            if ((lpWindow.componentexist(0) == 1) && (lpWindow.Visible == 1) && !StringIn(lpWindow.WindowState, "Minimized"))
            {
                lpWindow.Show("Minimize");
            }
            return result;
        }
        public virtual bool Restore() // Restore the Visual Studio
        {
            bool result = false;
            RPAWin32Component lpWindow = (RPAWin32Component)vsWindow.getComponentRoot();

            if ((lpWindow.componentexist(0) == 1) && StringIn(lpWindow.WindowState, "Minimized", "Maximized"))
            {
                lpWindow.Show("Restore");
            }
            return result;
        }
    }

    public abstract class MyCartesProcessBase : MyCartesProcess
    {
        private ClassVisualStudio fVS = null;
        
        public MyCartesProcessBase() : base("$Abort")
        {
        }

        protected override void MergeLibrariesAndLoadVariables()
        {
            try
            {
                cartes.merge(CurrentPath + "\\Cartes\\RPABaseAPI.cartes.rpa");
                base.MergeLibrariesAndLoadVariables();
            }
            catch (Exception e)
            {
                forensic("MyCartesProcessBase.MergeLibrariesAndLoadVariables", e);
                throw new MyException(EXIT_ERROR, e.Message);
            }
        }
        protected override bool DoInit()
        {
            VisualStudio.Minimice();
            return base.DoInit();
        }
        protected override void DoEnd() 
        {
            VisualStudio.Restore();
            base.DoEnd();
        }
        protected static DateTime GetFormatDateTime(string mask, string value)
        {
            CE_Data.DateTime dt = new CE_Data.DateTime();
            dt.Text[mask] = value;
            return new DateTime(int.Parse(dt.Text["yyyy"]), int.Parse(dt.Text["mm"]), int.Parse(dt.Text["dd"]),
                                int.Parse(dt.Text["hh"]), int.Parse(dt.Text["nn"]), int.Parse(dt.Text["ss"]));
        }
        protected virtual ClassVisualStudio GetVisualStudio()
        {
            if (fVS == null) fVS = new ClassVisualStudio(this);
            return fVS;
        }

        public ClassVisualStudio VisualStudio
        {
            get { return GetVisualStudio(); }
        } // Read
    }
}
