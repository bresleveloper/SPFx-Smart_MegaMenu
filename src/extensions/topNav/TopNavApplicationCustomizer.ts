import { override } from '@microsoft/decorators';
import { Guid, Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName, PlaceholderProvider
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import styles from './AppCustomizer.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';

import * as strings from 'TopNavApplicationCustomizerStrings';

const LOG_SOURCE: string = 'TopNavApplicationCustomizer';
const HEADER_TEXT: string = "This is the top zone";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ITopNavApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}


/** A Custom Action which can be run during execution of a Client Side Application */
export default class TopNavApplicationCustomizer
  extends BaseApplicationCustomizer<ITopNavApplicationCustomizerProperties> {

  @override
  protected onInit(): Promise<void> {/**replace func it to on get terms init init */
    console.log('SmartMegaMenu onInit 1.0')

    this.loadScripts().then(()=>{
      console.log('loadScripts finish (then)');




      this.getNavigationTerms().then(()=>{
        console.log('getNavigationTerms finish (then)');

      });
    })

    return super.onInit();
  }

  //protected onInit(): Promise<void> {
  //   window["megaTopNav"] = {
  //     getTermSet:this.getTermSet.bind(this),
  //     extention:this,
  //     getTerms:this.getTermSet.bind(this),
  //     getTermSetAsTree:this.getTermSetAsTree.bind(this)
  //   };
//
  //  let guid:SP.Guid = new SP.Guid('f408eaa2-df1e-4d84-af6a-9d256d21b8fa');
  //  this.getTermsInit(guid);
  //   Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
  //   /*now catch this and insert navigation*/
  //   let topPlaceholder: PlaceholderContent = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);
  //   if (topPlaceholder) {
  //     topPlaceholder.domElement.innerHTML = `<div class="${styles.app}">
  //                 <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.header}">
  //                   <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i>&nbsp; ${escape(HEADER_TEXT)}
  //                 </div>
  //               </div>`;
  //   }
////
////
  //   return Promise.resolve();
  //}

  loadScripts():Promise<void>{
    console.log('SmartMegaMenu - loadScripts')
    let siteColUrl= this.context.pageContext.site.absoluteUrl;
    return new Promise<void>((resolve_loadScripts, reject) => {

      SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/init.js', {
          globalExportsName: '$_global_init'
        })
          .then((): Promise<{}> => {
            return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/MicrosoftAjax.js', {
              globalExportsName: 'Sys'
            });
          })
          .then((): Promise<{}> => {
            return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.Runtime.js', {
              globalExportsName: 'SP'
            });
          })
          .then((): Promise<{}> => {
            return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.js', {
              globalExportsName: 'SP'
            });
          })
          .then((): Promise<{}> => {
            return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.publishing.js', {
              globalExportsName: 'SP'
            });
          })          
          .then((): Promise<{}> => {
            return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.requestexecutor.js', {
              globalExportsName: 'SP'
            });
          })
          .then((): Promise<{}> => {
            return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.taxonomy.js', {
              globalExportsName: 'SP'
            });
          })
          .then(():void => resolve_loadScripts());
    });
  }


  getNavigationTerms():Promise<{}>{
    console.log('SmartMegaMenu - getNavigationTerms')
    let myPromise = new Promise<{}>((resolve, reject) => {

      let siteColUrl = this.context.pageContext.site.absoluteUrl;
      let context: SP.ClientContext = new SP.ClientContext(siteColUrl);
      let factory = new SP.ProxyWebRequestExecutorFactory(siteColUrl);//appWebURL
      context.set_webRequestExecutorFactory(factory);
      let appContextSite = new SP.AppContextSite(context, siteColUrl);// $('#txtSharePointUrl').val());
      let hostWeb = appContextSite.get_web();
      let taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);


      let webNavSettings = new SP.Publishing.Navigation.WebNavigationSettings(context, hostWeb);
      context.load(webNavSettings.get_currentNavigation())
      let currentNavigationSettings = webNavSettings.get_currentNavigation();
      context.load(currentNavigationSettings);

      context.executeQueryAsync( ()=> {

        console.log('SmartMegaMenu - getNavigationTerms after query')

        let termStoreId = currentNavigationSettings.get_termStoreId();
        let termSetId = currentNavigationSettings.get_termSetId();
        let termStore = taxonomySession.get_termStores().getById(termStoreId);
        let termSet = termStore.getTermSet(termSetId);

        let navTermSet = SP.Publishing.Navigation.NavigationTermSet.getAsResolvedByWeb(context, termSet, hostWeb, 'CurrentNavigationSwitchableProvider');
        context.load(navTermSet);

        navTermSet = SP.Publishing.Navigation.NavigationTermSet.getAsResolvedByWeb(context, termSet, context.get_web(), 'GlobalNavigationTaxonomyProvider');

        let terms = navTermSet.get_terms();
        context.load(terms, 'Include(Id, Title, TargetUrl, FriendlyUrlSegment, Terms)');

        let allTerms = navTermSet.getAllTerms();
        context.load(allTerms, 'Include(Id, Title, TargetUrl, FriendlyUrlSegment, Terms)');

        context.executeQueryAsync(function (sender, args) {
          console.log('SmartMegaMenu - getNavigationTerms after nav query')

          let termsArr = [];
          console.log('SmartMegaMenu - getNavigationTerms allTerms : ',allTerms);
          let termsEnum = terms.getEnumerator();
          console.log('SmartMegaMenu loop navTerms');
  
          while (termsEnum.moveNext()) {
            //for the current term stub, get all the properties from the fully loaded getAllTerms object
            let currentTerm = termsEnum.get_current()
            console.log('SmartMegaMenu currentTerm', currentTerm);
            window['x'] = currentTerm;
  
            let newTerm = {
              "id": currentTerm.get_id().toString(),
              "name": currentTerm.get_title().get_value(),
              "href": currentTerm.get_targetUrl().get_value(),
              "childnodes": []
            }

            termsArr.push(newTerm)
          }
        });

      });
    });

    return myPromise;
  }


}
