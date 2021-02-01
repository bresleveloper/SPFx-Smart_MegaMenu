import styles from './AppCustomizer.module.scss';
import { override } from '@microsoft/decorators';
import { Guid, Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName, PlaceholderProvider
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';


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
  public ajaxCounter:number = 0;
  public MegaMenuListData:[] = [];
  public TSguid:string = '';
  public pageDirection:string = '';
  public masterTreesDictionary:{} = {};
  public trees = [];
  public settings:{}={};

  @override
  protected onInit(): Promise<void> {/**replace func it to on get terms init init */
    
    window['MegaMenuInfo']={
      MegaMenuListData:this.MegaMenuListData,
    };
    console.log('SmartMegaMenu onInit 1.0')

    this.getSettings().then(()=>{
      console.log('getSettings resolved');

      this.loadScripts().then(()=>{
        console.log('loadScripts finish (then)');

        this.getTermSetAsTree().then((tree)=>{
          window['tree'] = tree;
          console.log('getTermSetAsTree finish (then)', tree);
          //switch mega-type
          this.buildHtmlBlue(tree);
        });
      })
    }, /*  */ ()=>{
      console.log('getSettings reject');
    })
    return super.onInit();
  }

  public getSettings():Promise<void>{
    return new Promise<void>((resolve, reject)=>{
      let listname:string = 'MegaMenuSettings';
      this.getListItems(listname).then(()=>{
        console.log('this.MegaMenuListData : ',this.MegaMenuListData);
        window['mmList']=this.MegaMenuListData;
        for(let i=0;i<this.MegaMenuListData.length;i++){
          const item = this.MegaMenuListData[i];
          if(i==0){
            this.TSguid = item['Value'];
            this.TSguid = this.TSguid.trim();
            console.log('this.TSguid '+this.TSguid);
          }
          if(i==4 && item['Value']!=null){
            this.pageDirection = item['Value'];
            console.log('this.pageDirection '+this.pageDirection);
          }

        }
        resolve();
      });
      resolve();
      //list not ok
      reject();
      console.error("i dont have a guid")
    });
  }

  public buildHtmlBlue(tree:{}){
      //   /*now catch this and insert navigation*/
    let bench = tree["children"];
    window['bench']=bench;
    let inner:any = ``;
    let thLevel:any = ``;
    let fLeveltemplate = `<span class="ms-HorizontalNavItem ${styles.topNavSpan}" data-automationid="HorizontalNav-link">
                      <a class="#CLASS# ms-HorizontalNavItem-link is-not-selected ${styles.PermanentA}" href=#HREF# onmouseover="">
                        #NAME#
                      </a>
                    </span>`;
    let sLevelTemplate = `<ul><a class=${styles.sLevelA} href="#HREF#"><li class=${styles.sLevel}>#NAME#</li></a>#MORE#</ul>`;
    let thLevelTemplate = `<a class=${styles.thLevelA} href="#HREF#"><li class=${styles.thLevel}>#NAME#</li></a>`;
    for(let i=0;i<bench.length;i++){
      inner+=`<div class=${styles.wrapDiv} onmouseover="this.lastChild.style.visibility='visible';" onmouseleave="this.lastChild.style.visibility='hidden';">`;
      if(bench[i].localCustomProperties["_Sys_Nav_SimpleLinkUrl"]){
          inner+=fLeveltemplate.replace("#HREF#",bench[i].localCustomProperties["_Sys_Nav_SimpleLinkUrl"]).replace("#NAME#",bench[i].title);
      }
      else{
        if(bench[i].url){
          inner+=fLeveltemplate.replace("#HREF#",bench[i].url).replace("#NAME#",bench[i].title);
        }
        else{
          inner+=fLeveltemplate.replace("#NAME#",bench[i].title).replace("#HREF#","#");
        }
      }
      inner+=`<div class ="${styles.openDiv}" class="${String(i)}">`;
      for(let j=0;j<bench[i]['children'].length;j++){
        let subench = bench[i]['children'][j];
        let sHref = subench.localCustomProperties["_Sys_Nav_SimpleLinkUrl"]?subench.localCustomProperties["_Sys_Nav_SimpleLinkUrl"]:subench.url?subench.url:"#";
        if(subench['children']){
          for(let h=0;h<subench['children'].length;h++){
            let subsubench = subench['children'][h];
            let tHref = subsubench.localCustomProperties["_Sys_Nav_SimpleLinkUrl"]?subsubench.localCustomProperties["_Sys_Nav_SimpleLinkUrl"]:subsubench.url?subsubench.url:"#";
            thLevel+=thLevelTemplate.replace("#NAME#",subsubench.title).replace("#HREF#",tHref);
          }
        }
        inner+=sLevelTemplate.replace("#NAME#",subench.title).replace("#MORE#",thLevel).replace("#HREF#",sHref);
        thLevel=``;
      }
        inner+=`</div>`;
      inner+=`</div>`;
    }


     let topPlaceholder: PlaceholderContent = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);
     if (topPlaceholder) {
       topPlaceholder.domElement.innerHTML = `<div class="${styles.app}">
                   <div class= "${styles.header}">
                    ${inner}
                   </div>
                 </div>`;
      }
    }

  public loadScripts():Promise<void>{
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

  public getTerms():Promise<SP.Taxonomy.TermCollection>{
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

  public getTermSetAsTree():Promise<{}> {
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
                        term.url = term.localCustomProperties.url;
                      }
                      if (term.localCustomProperties && term.localCustomProperties._Sys_Nav_SimpleLinkUrl) {
                        term.url = term.localCustomProperties._Sys_Nav_SimpleLinkUrl;
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
  });
}

  public sortTermsFromTree(tree):any {
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
  public getListItems(listname:string): Promise<void> {
    let myPromise = new Promise<void>((resolve) => {
    console.log('asking list items for MegaMenuSettings');
    this.ajaxCounter++;

    this.context.spHttpClient.get(
      this.context.pageContext.site.absoluteUrl +
      `/_api/web/lists/GetByTitle('${listname}')/Items`, SPHttpClient.configurations.v1)
      //`/_api/web/lists/GetByTitle('${listname}')/Items`, SPHttpClient.configurations.v1)
          .then((response: SPHttpClientResponse) => {
              response.json().then((data)=> {

                  console.log('list items for', listname, data);
                  this.ajaxCounter--;
                  this.MegaMenuListData = data.value;
                  if (this.ajaxCounter == 0) {
                    resolve();
                  }
                  else{
                    console.error('could not get list')
                  }

              });
          });
    });
    return myPromise;
  }


}
