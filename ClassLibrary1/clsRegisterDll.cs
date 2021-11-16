using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ClassLibrary1
{
    [RunInstaller(true)]
    public partial class clsRegisterDll : System.Configuration.Install.Installer
    {
        public clsRegisterDll()
        {
            InitializeComponent();
        }

        [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Demand)]
        public override void Install(IDictionary stateSaver)
        {
            base.Install(stateSaver);
            System.Diagnostics.Process.Start("C:\\Program Files (x86)\\COM_CADLib1C\\COM-component\\Reg.bat");
            //RegistrationServices regsrv = new RegistrationServices();
            //if (!regsrv.RegisterAssembly(GetType().Assembly, AssemblyRegistrationFlags.SetCodeBase))
            //{
            //    throw new InstallException("Failed to register for COM Interop.");
            //}
        }

        [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Demand)]
        public override void Uninstall(IDictionary savedState)
        {
            System.Diagnostics.Process.Start("C:\\Program Files (x86)\\COM_CADLib1C\\COM-component\\UnReg.bat");
            base.Uninstall(savedState);
            //RegistrationServices regsrv = new RegistrationServices();
            //if (!regsrv.UnregisterAssembly(GetType().Assembly))
            //{
            //    throw new InstallException("Failed to unregister for COM Interop.");
            //}
        }

    }
}
