﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace To.Rpa.AppCapital.Interfaces
{
    public interface IUserInterface
    {
        void CheckCompleteLoad();
        void DoActivities();
        void DoOperations();
        void DoUIValidation();
        void ExtractElements();
    }
}
