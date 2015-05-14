using System;
using System.Linq;
using IonFar.SharePoint.Provisioning.Infrastructure;
using Microsoft.SharePoint.Client;

namespace IonFar.SharePoint.Provisioning.Services
{
    public class GroupUpdateInformation
    {
        public GroupUpdateInformation(string oldTitle, string newTitle)
        {
            OldTitle = oldTitle;
            NewTitle = newTitle;
        }

        public string OldTitle { get; private set; }
        public string NewTitle { get; private set; }
    }


    public class GroupDeletionInformation
    {
        public GroupDeletionInformation(string title)
        {
            Title = title;
            
        }

        public string Title { get; private set; }
       
    }
    public interface ISecurityGroupProvisioningService
    {
        void CreateSecurityGroups(GroupCreationInformation[] groupCreationInformations);
        void UpdateSecurityGroups(GroupUpdateInformation[] groupUpdateInformations);
        void DeleteSecurityGroups(GroupDeletionInformation[] groupDeletionInformations);
    }

    public class SecurityGroupProvisioningService : ISecurityGroupProvisioningService
    {
        private readonly ClientContext _clientContext;
        private readonly IProvisionLog _logger;

        public SecurityGroupProvisioningService(ClientContext clientContext, IProvisionLog logger)
        {
            _clientContext = clientContext;
            _logger = logger;
        }

        public void CreateSecurityGroups(GroupCreationInformation[] groupCreationInformations)
        {
            var web = _clientContext.Web;
            var groups = web.SiteGroups;
            _clientContext.Load(groups);
            _clientContext.ExecuteQuery();

            foreach (var groupCreationInformation in groupCreationInformations)
            {
                if (groups.Any(g => g.Title == groupCreationInformation.Title))
                {
                    _logger.Information("Group {0} already exists, continuing.", groupCreationInformation.Title);
                    continue;
                }

                _logger.Information("Creating Security Group {0}", groupCreationInformation.Title);
                web.SiteGroups.Add(groupCreationInformation);
            }

            _clientContext.ExecuteQuery();
        }

        public void UpdateSecurityGroups(GroupUpdateInformation[] groupUpdateInformations)
        {
            try
            {
                var web = _clientContext.Web;
                var groups = web.SiteGroups;
                _clientContext.Load(groups);
                _clientContext.ExecuteQuery();

                foreach (var gui in groupUpdateInformations)
                {
                    if (groups.Any(g => g.Title == gui.OldTitle))
                    {
                        _logger.Information("Found existing group '{0}' will try to update its Title to {1}.",
                            gui.OldTitle,
                            gui.NewTitle);

                        var group = groups.GetByName(gui.OldTitle);
                        if (group != null)
                        {
                            group.Title = gui.NewTitle;
                            group.Update();
                            _logger.Information("Updated group title to '{0}'.", gui.NewTitle);
                        }
                        else
                        {
                            _logger.Warning("Call to groups.GetByName(gui.OldTitle) failed for '{0}'.", gui.OldTitle);
                        }
                    }
                    else
                    {
                        _logger.Warning("Could not locate existing group '{0}'.", gui.OldTitle);
                    }
                }

                _clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                _logger.Error("UpdateSecurityGroups(GroupUpdateInformation[] groupUpdateInformations) FAILED! Exception: {0}", ex);
            }
        }

        public void DeleteSecurityGroups(GroupDeletionInformation[] groupDeletionInformations)
        {
            try
            {
                var web = _clientContext.Web;
                var groups = web.SiteGroups;
                _clientContext.Load(groups);
                _clientContext.ExecuteQuery();


                
                foreach (var gdi in groupDeletionInformations)
                {
                    if (groups.Any(g => g.Title == gdi.Title))
                    {
                        _logger.Information("Found existing group '{0}' will try to delete it.",
                            gdi.Title
                            );

                        var group = groups.GetByName(gdi.Title);
                        if (group != null)
                        {
                            groups.Remove(group);
                            
                            _logger.Information("Deleted group '{0}'.", gdi.Title);
                        }
                        else
                        {
                            _logger.Warning("Call to groups.GetByName(gdi.Title) failed for '{0}'.", gdi.Title);
                        }
                    }
                    else
                    {
                        _logger.Warning("Could not locate existing group '{0}'.", gdi.Title);
                    }
                }

                _clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                _logger.Error("UpdateSecurityGroups(GroupUpdateInformation[] groupUpdateInformations) FAILED! Exception: {0}", ex);
            }
        }


       
    }
}
