﻿using System;
namespace IonFar.SharePoint.Provisioning.Infrastructure
{
    public interface ILogger
    {
        void Information(string message, params object[] args);
        void Warning(string message, params object[] args);
        void Error(Exception ex, string message, params object[] args);
    }
}
