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

  public getTermSet(){
    let siteColUrl= this.context.pageContext.site.absoluteUrl;

  }

  public getTermSetAsTree(terms: SP.Taxonomy.TermCollection){

  }




  @override
  protected onInit(): Promise<void> {/**replace func it to on get terms init init */
    console.log('on getTermsInit')
    let siteColUrl= this.context.pageContext.site.absoluteUrl;
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
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.taxonomy.js', {
            globalExportsName: 'SP'
          });
        })
        .then((): void => {
            let spContext: SP.ClientContext =new SP.ClientContext(siteColUrl);
            let taxSession =  SP.Taxonomy.TaxonomySession.getTaxonomySession(spContext);
            console.log("taxonomy session ", taxSession)
            let termStore  = taxSession.getDefaultSiteCollectionTermStore();

            let guid:SP.Guid = new SP.Guid('f408eaa2-df1e-4d84-af6a-9d256d21b8fa');
            let termSet = termStore.getTermSet(guid);
            let terms = termSet.getAllTerms();
            console.log('this final terms',terms);
            console.log('send to tree');
            //this.getTermSetAsTree(terms);
            spContext.load(terms);
            spContext.executeQueryAsync( ()=> {
            var termsEnum = terms.getEnumerator();
            let termDepartment:any[]=[];
            while (termsEnum.moveNext()) {
              var spTerm = termsEnum.get_current();
              termDepartment.push({label:spTerm.get_name(),value:spTerm.get_name(), id:spTerm.get_id()});
            }

             window['termDepartment']= termDepartment;
             console.log('window[tDep] ', window['termDepartment']);
            });

          });

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
}
