using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration.Extensibility;

namespace TestRibbonLoad
{
    [ComVisible(true)]
    public class MyRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            Debug.WriteLine("MyRibbon.GetCustomUI");
            return
@"<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
    <ribbon>
    <tabs>
        <tab id='CustomTab' label='Test Ribbon'>
            <group id='SampleGroup' label='Buttons'>
                <button id='Button1' label='Press Me!' size='large' onAction='OnButtonPress' />
            </group >
        </tab>
    </tabs>
    </ribbon>
</customUI>
";
        }

        public void OnButtonPress(IRibbonControl control)
        {
            MessageBox.Show("Hello from TestRibbonLoad add-in");
        }

        public override void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            Debug.WriteLine("MyRibbon.OnConnection");
            base.OnConnection(Application, ConnectMode, AddInInst, ref custom);
        }

        public override void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            base.OnDisconnection(RemoveMode, ref custom);
        }

        public override void OnAddInsUpdate(ref Array custom)
        {
            Debug.WriteLine("MyRibbon.OnConnection");
            base.OnAddInsUpdate(ref custom);
        }
    }
}
