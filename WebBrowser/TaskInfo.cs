using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TenDayBrowser
{
    using System;
    using System.Runtime.InteropServices;
    public enum TaskCommand
    {
        Task_None = -1,
        Task_Wait = 0,

        Task_InputText = 1,
        Task_ClickButton = 2,
        Task_ClickLink = 3,
        Task_Navigate = 4,
        Task_DeepClick = 5,
        Task_ClearCookie = 6,
        Task_FindLinkLinkPage1 = 7,
        Task_FindLinkLinkPage2 = 8,
        Task_Fresh = 9,
        Task_PressKey = 10,
        Task_ClickRadio = 11,
        Task_ClickChecked = 12,

        Task_FindLinkHrefPage1 = 13,
        Task_FindLinkHrefPage2 = 14,
        Task_FindHrefLinkPage1 = 15,

        Task_FindHrefLinkPage2 = 16,//16
        Task_FindHrefHrefPage1 = 17,
        Task_FindHrefHrefPage2 = 18,
        Task_FindLinkSrcPage1 = 19,
        Task_FindLinkSrcPage2 = 20,
        Task_FindHrefSrcPage1 = 21,
        Task_FindHrefSrcPage2 = 22,
        Task_FindSrcLinkPage1 = 23,
        Task_FindSrcLinkPage2 = 24,
        Task_FindSrcHrefPage1 = 25,
        Task_FindSrcHrefPage2 = 26,
        Task_FindSrcSrcPage1 = 27,
        Task_FindSrcSrcPage2 = 28,
        Task_GoPage = 29,
        Task_VisitCompare = 30,
        Task_ClickCompare = 31,
        Task_ClickMe,
        Task_VisitPage,

        Task_Count,
        
    }

    public class TaskInfo
    {
        public string _param1;
        public string _param2;
        public string _param3;
        public string _param4;
        public string _param5;

        public TaskInfo(string param1,  [Optional, DefaultParameterValue("")]　string param2,  [Optional, DefaultParameterValue("")]　string param3,  [Optional, DefaultParameterValue("bb")] string param4,  [Optional, DefaultParameterValue("bb")]　string param5)
        {
            this._param1 = param1;
            this._param2 = param2;
            this._param3 = param3;
            this._param4 = param4;
            this._param5 = param5;
        }

        public uint CalculateScore()
        {
            uint num = 0;
            switch (((TaskCommand) WindowUtil.StringToInt(this._param1)))
            {
                case TaskCommand.Task_Wait:
                    if (!string.IsNullOrEmpty(this._param2))
                    {
                        num = (WindowUtil.StringToUint(this._param2) + 0x1d) / 30;
                    }
                    return num;

                case TaskCommand.Task_DeepClick:
                    if (!string.IsNullOrEmpty(this._param2))
                    {
                        num = 1 + (WindowUtil.StringToUint(this._param2) * ((WindowUtil.StringToUint(this._param3) + 0x1d) / 30));
                    }
                    return num;

                case TaskCommand.Task_FindLinkLinkPage1:
                case TaskCommand.Task_FindLinkHrefPage1:
                case TaskCommand.Task_FindHrefLinkPage1:
                case TaskCommand.Task_FindHrefHrefPage1:
                case TaskCommand.Task_FindSrcLinkPage1:
                case TaskCommand.Task_FindSrcHrefPage1:
                case TaskCommand.Task_FindHrefSrcPage1:
                case TaskCommand.Task_FindLinkSrcPage1:
                case TaskCommand.Task_FindSrcSrcPage1:
                    if (!string.IsNullOrEmpty(this._param4))
                    {
                        num = WindowUtil.StringToUint(this._param4) + 1;
                    }
                    return num;
            }
            return 1;
        }
    }
}
