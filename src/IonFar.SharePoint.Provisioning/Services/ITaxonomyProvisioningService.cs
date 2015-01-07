using System;
using IonFar.SharePoint.Provisioning.Services.Taxonomy;

namespace IonFar.SharePoint.Provisioning.Services
{
    public interface ITaxonomyProvisioningService
    {
        Guid DefaultTermStoreId { get; }
        void AddGroup(string groupName, Guid groupId);
        void AddNavigationTerms(NavigationTermInfo[] navigationTerms, Guid termSetId);
        void AddNavigationTermsToTerm(NavigationTermInfo[] navigationTerms, Guid termId);
        void AddTermSet(string termSetName, Guid termSetId, Guid groupId, bool isOpen, bool isNavigationTermSet = false);
        void AddTerms(string[] termNames, Guid termSetId);
        void AddTerms(TermInfo[] termInfos, Guid termSetId);
        void AddTermsToTerm(string[] termNames, Guid termId);
        void DeleteGroup(Guid groupId);
        void DeleteTermSet(Guid termSetId);
    }
}