using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TalentClassLibrary
{
    public delegate void ClickEventHandler(object sender, EventArgs e);

    public class DelProjectExperienceBtn
    {
        /// <summary>
        /// 宣告Botton的事件
        /// </summary>
        public event ClickEventHandler ClickEvent;

        /// <summary>
        /// 建立Button成員方法
        /// </summary>
        public void Click()
        {
            if (ClickEvent != null)
            {
                MessageBox.Show("事件開始");
                //拋出事件，給所有相應者
                ClickEvent(this, null);
            }
        }
    }
}
