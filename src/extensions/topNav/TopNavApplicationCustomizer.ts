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

  //guid:SP.Guid = new SP.Guid('f408eaa2-df1e-4d84-af6a-9d256d21b8fa');
  TSguid:string = 'f408eaa2-df1e-4d84-af6a-9d256d21b8fa';
  masterTreesDictionary:{} = {};
  trees = [];
  settings:{}={};

  @override
  protected onInit(): Promise<void> {/**replace func it to on get terms init init */
    console.log('SmartMegaMenu onInit 1.0')

    //this.getSettings().then(()=>{

    this.loadScripts().then(()=>{
      console.log('loadScripts finish (then)');

      this.getTermSetAsTree().then((tree)=>{
        console.log('getTermSetAsTree finish (then)', tree);

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

  getTerms():Promise<SP.Taxonomy.TermCollection>{
    console.log('SmartMegaMenu - getNavigationTerms')
    let myPromise = new Promise<SP.Taxonomy.TermCollection>((resolve, reject) => {

      let siteColUrl= this.context.pageContext.site.absoluteUrl;
      let spContext: SP.ClientContext = new SP.ClientContext(siteColUrl);
      let taxSession =  SP.Taxonomy.TaxonomySession.getTaxonomySession(spContext);
      let termStore  = taxSession.getDefaultSiteCollectionTermStore();
      let guid:SP.Guid = new SP.Guid(this.TSguid);
      let termSet = termStore.getTermSet(guid);
      let terms = termSet.getAllTerms();
      spContext.load(terms);
      spContext.executeQueryAsync( ()=> {
        resolve(terms);
      })
    });
    return myPromise;
  }

  getTermSetAsTree():Promise<{}> {
    return new Promise<{}>((resolve) => {
      this.getTerms().then( (terms) => {
          let termsEnumerator = terms.getEnumerator(),
              tree = {
                  term: terms,
                  children: []
              };
          //ariel
          let termsDict = {}

          // Loop through each term
          while (termsEnumerator.moveNext()) {
              let currentTerm = termsEnumerator.get_current();
              let currentTermPath = currentTerm.get_pathOfTerm().split(';');
              let children = tree.children;

              // Loop through each part of the path
              for (let i = 0; i < currentTermPath.length; i++) {
                  let foundNode = false;

                  let j;
                  for (j = 0; j < children.length; j++) {
                      if (children[j].name === currentTermPath[i]) {
                          foundNode = true;
                          break;
                      }
                  }

                  // Select the node, otherwise create a new one
                  let term = foundNode ? children[j] : { name: currentTermPath[i], children: [] };

                  // If we're a child element, add the term properties
                  if (i === currentTermPath.length - 1) {
                      term.term = currentTerm;
                      term.title = currentTerm.get_name();
                      term.guid = currentTerm.get_id().toString();
                      term.description = currentTerm.get_description();
                      term.customProperties = currentTerm.get_customProperties();
                      term.localCustomProperties = currentTerm.get_localCustomProperties();

                      term.url = '';
                      if (term.localCustomProperties && term.localCustomProperties.url) {
                        term.url = term.localCustomProperties.url
                      }
                      if (term.localCustomProperties && term.localCustomProperties._Sys_Nav_SimpleLinkUrl) {
                        term.url = term.localCustomProperties._Sys_Nav_SimpleLinkUrl
                      }

                  }

                  //ariel
                  termsDict[term.guid] = term;
                  this.masterTreesDictionary[term.guid] = term;

                  // If the node did exist, let's look there next iteration
                  if (foundNode) {
                      children = term.children;
                  }
                  // If the segment of path does not exist, create it
                  else {
                      children.push(term);

                      // Reset the children pointer to add there next iteration
                      if (i !== currentTermPath.length - 1) {
                          children = term.children;
                      }
                  }
              }
          }

          tree = this.sortTermsFromTree(tree);
          this.trees.push({tree:tree, dict:termsDict})

          resolve(tree);
      });
  })
}

  sortTermsFromTree(tree):any {
    // Check to see if the get_customSortOrder function is defined. If the term is actually a term collection,
    // there is nothing to sort.
    if (tree.children.length && tree.term.get_customSortOrder) {
        let sortOrder = null;

        if (tree.term.get_customSortOrder()) {
            sortOrder = tree.term.get_customSortOrder();
        }

        // If not null, the custom sort order is a string of GUIDs, delimited by a :
        if (sortOrder) {
            sortOrder = sortOrder.split(':');

            tree.children.sort(function (a, b) {
                let indexA = sortOrder.indexOf(a.guid);
                let indexB = sortOrder.indexOf(b.guid);

                if (indexA > indexB) {
                    return 1;
                } else if (indexA < indexB) {
                    return -1;
                }

                return 0;
            });
        }
        // If null, terms are just sorted alphabetically
        else {
            tree.children.sort(function (a, b) {
                if (a.title > b.title) {
                    return 1;
                } else if (a.title < b.title) {
                    return -1;
                }

                return 0;
            });
        }
    }

    for (let i = 0; i < tree.children.length; i++) {
        tree.children[i] = this.sortTermsFromTree(tree.children[i]);
    }

    return tree;
  }


}
