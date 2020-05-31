using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace getJVLInkData
{
    public class OperateForm
    {
        Form1 _form1;
        public OperateForm(Form1 form1)
        {
            _form1 = form1;
        }
        public void enableButton()
        {
            _form1.button1.Enabled = true;
            _form1.dateTimePicker1.Enabled = true;
            _form1.btnGetJVData.Enabled = true;
            _form1.button4.Enabled = true;
            _form1.button5.Enabled = true;
        }

        public void disableButton()
        {
            _form1.button1.Enabled = false;
            _form1.dateTimePicker1.Enabled = false;
            _form1.btnGetJVData.Enabled = false;
            _form1.button4.Enabled = false;
            _form1.button5.Enabled = false;
        }
    }
}
