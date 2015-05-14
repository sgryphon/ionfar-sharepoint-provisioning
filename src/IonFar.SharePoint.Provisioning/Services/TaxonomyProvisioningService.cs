using System;
using IonFar.SharePoint.Provisioning.Infrastructure;
using IonFar.SharePoint.Provisioning.Services.Taxonomy;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing.Navigation;
using Microsoft.SharePoint.Client.Taxonomy;

namespace IonFar.SharePoint.Provisioning.Services
{
    public class TaxonomyProvisioningService : ITaxonomyProvisioningService
    {
        private readonly ClientContext _clientContext;
        private readonly IProvisionLog _logger;
        private readonly TaxonomySession _taxonomySession;
        private const int DefaultLcid = 1033;
        public Guid DefaultTermStoreId { get; private set; }

        public TaxonomyProvisioningService(ClientContext clientContext, IProvisionLog logger)
        {
            _clientContext = clientContext;
            _logger = logger;
            _taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);
            LoadDefaultTermStoreId();
        }

        public void AddGroup(string groupName, Guid groupId)
        {
            _logger.Information("Adding group '{0}' with Id {1}", groupName, groupId);

            var termStore = _taxonomySession.TermStores.GetById(DefaultTermStoreId);
            var termGroup = termStore.CreateGroup(groupName, groupId);

            _clientContext.Load(termGroup);
            termStore.CommitAll();

            _clientContext.ExecuteQuery();
        }

        public void AddTermSet(string termSetName, Guid termSetId, Guid groupId, bool isOpen = false, bool isNavigationTermSet = false)
        {
            _logger.Information("Adding TermSet '{0}' with Id {1} to GroupId {2}", termSetName, termSetId, groupId);

            var termStore = _taxonomySession.TermStores.GetById(DefaultTermStoreId);
            var termGroup = termStore.GetGroup(groupId);
            var termSet = termGroup.CreateTermSet(termSetName, termSetId, DefaultLcid);

            _clientContext.Load(termSet);
            termSet.IsOpenForTermCreation = isOpen;
            termStore.CommitAll();

            _clientContext.ExecuteQuery();

            if (isNavigationTermSet)
            {
                NavigationTermSet navTermSet = NavigationTermSet.GetAsResolvedByWeb(
                    _clientContext,
                    termSet, _clientContext.Web, "GlobalNavigationTaxonomyProvider");

                navTermSet.IsNavigationTermSet = true;

                termStore.CommitAll();
                _clientContext.ExecuteQuery();
            }
        }

        public void AddNavigationTerms(NavigationTermInfo[] navigationTerms, Guid termSetId)
        {
            _logger.Information("Adding Navigation Terms to TermSetId '{0}'", termSetId);

            var termStore = _taxonomySession.TermStores.GetById(DefaultTermStoreId);
            var termSet = termStore.GetTermSet(termSetId);
            var navTermSet = NavigationTermSet.GetAsResolvedByWeb(_clientContext,
            termSet, _clientContext.Web, "GlobalNavigationTaxonomyProvider");

            _clientContext.Load(termSet);
            foreach (var termInfo in navigationTerms)
            {
                _logger.Information("Creating term '{0}'", termInfo.Name);
                var navigationTermGuid = termInfo.TermId == Guid.Empty ? Guid.NewGuid() : termInfo.TermId;
                var navigationTerm = navTermSet.CreateTerm(termInfo.Name, NavigationLinkType.SimpleLink, navigationTermGuid);
                navigationTerm.SimpleLinkUrl = termInfo.Url;
            }

            termStore.CommitAll();

            _clientContext.ExecuteQuery();
        }

        public void AddNavigationTermsToTerm(NavigationTermInfo[] navigationTerms, Guid termId)
        {
            _logger.Information("Adding Navigation Terms to TermSetId '{0}'", termId);

            var termStore = _taxonomySession.TermStores.GetById(DefaultTermStoreId);
            var parentTerm = _taxonomySession.GetTerm(termId);
            _clientContext.Load(parentTerm);

            var parentNavigationterm = NavigationTerm.GetAsResolvedByWeb(_clientContext, parentTerm,
                _clientContext.Web, "GlobalNavigationTaxonomyProvider");

            foreach (var termInfo in navigationTerms)
            {
                _logger.Information("Creating term '{0}'", termInfo.Name);
                var navigationTermGuid = termInfo.TermId == Guid.Empty ? Guid.NewGuid() : termInfo.TermId;
                var navigationTerm = parentNavigationterm.CreateTerm(termInfo.Name, NavigationLinkType.SimpleLink, navigationTermGuid);
                navigationTerm.SimpleLinkUrl = termInfo.Url;
            }

            termStore.CommitAll();

            _clientContext.ExecuteQuery();
        }

        public void AddTerms(string[] termNames, Guid termSetId)
        {
            _logger.Information("Adding Terms to TermSetId '{0}'", termSetId);

            var termStore = _taxonomySession.TermStores.GetById(DefaultTermStoreId);
            var termSet = termStore.GetTermSet(termSetId);

            _clientContext.Load(termSet);
            foreach (string termName in termNames)
            {
                _logger.Information("Creating term '{0}'", termName);
                termSet.CreateTerm(termName, DefaultLcid, Guid.NewGuid());
            }

            termStore.CommitAll();

            _clientContext.ExecuteQuery();
        }

        public void AddTerms(TermInfo[] termNames, Guid termSetId)
        {
            _logger.Information("Adding Terms to TermSetId '{0}'", termSetId);

            var termStore = _taxonomySession.TermStores.GetById(DefaultTermStoreId);
            var termSet = termStore.GetTermSet(termSetId);

            _clientContext.Load(termSet);
            foreach (var termInfo in termNames)
            {
                var termId = termInfo.TermId == Guid.Empty ? Guid.NewGuid() : termInfo.TermId;
                _logger.Information("Creating term '{0}'", termInfo.Name);
                termSet.CreateTerm(termInfo.Name, DefaultLcid, termId);
            }

            termStore.CommitAll();

            _clientContext.ExecuteQuery();
        }

        public void AddTermsToTerm(string[] termNames, Guid termId)
        {
            _logger.Information("Adding Terms to Term '{0}'", termId);

            var termStore = _taxonomySession.TermStores.GetById(DefaultTermStoreId);
            var parentTerm = _taxonomySession.GetTerm(termId);
            _clientContext.Load(parentTerm);

            foreach (string termName in termNames)
            {
                _logger.Information("Creating term '{0}'", termName);
                parentTerm.CreateTerm(termName, DefaultLcid, Guid.NewGuid());
            }

            termStore.CommitAll();

            _clientContext.ExecuteQuery();
        }

        public void DeleteGroup(Guid groupId)
        {
            _logger.Warning("Deleting GroupId {0}", groupId);

            var termStore = _taxonomySession.TermStores.GetById(DefaultTermStoreId);
            var groupToDelete = termStore.GetGroup(groupId);

            groupToDelete.DeleteObject();

            _clientContext.ExecuteQuery();
        }

        public void DeleteTermSet(Guid termSetId)
        {
            _logger.Warning("Deleting TermSetId '{0}'", termSetId);

            var termStore = _taxonomySession.TermStores.GetById(DefaultTermStoreId);
            var termSetToDelete = termStore.GetTermSet(termSetId);

            termSetToDelete.DeleteObject();

            _clientContext.ExecuteQuery();
        }

        private void LoadDefaultTermStoreId()
        {
            var termStore = _taxonomySession.GetDefaultSiteCollectionTermStore();
            
            _clientContext.Load(termStore, t => t.Id);

            _clientContext.ExecuteQuery();

            DefaultTermStoreId = termStore.Id;
            _logger.Information("Default term store id is: '{0}'", DefaultTermStoreId);
        }
    }
}
