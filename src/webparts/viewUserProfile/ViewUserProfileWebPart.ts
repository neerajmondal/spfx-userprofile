import { Version } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ViewUserProfileWebPart.module.scss';
import * as strings from 'ViewUserProfileWebPartStrings';
import { UserObjectType } from './UserTypes';
import * as $ from 'jquery';
require('bootstrap') ;
require('./basestyles.css');



export interface IViewUserProfileWebPartProps {
  description: string;
  Department:string;
}
export {UserObjectType} from './UserTypes';

export default class ViewUserProfileWebPart extends BaseClientSideWebPart<IViewUserProfileWebPartProps> {

  private _listOptions: IPropertyPaneDropdownOption[] = [];
  private selectedSiteColKey:string = '0';

  private loadIndicator : boolean = true;
  private siteColUrl:string;

  constructor()
  {
    super();

    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css');
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css');
 

  }
  protected onInit(): Promise<void> {


    return super.onInit();
  }



  public render(): void {

    if (!this.renderedOnce) {
    this.domElement.innerHTML = `
    <section class="team pb-5">
          <div class="container">
              <h4 class="section-title"> Department</h4>
              
              <div id="row-container" class="row">


                  </div>
          </div>
      </section>
<!-- Team -->

`;
    }

let strdepartment:string="";

if(this.properties.Department!= undefined){
    strdepartment=this.properties.Department;
}


if(strdepartment.length>0){

let url=this.context.pageContext.web.absoluteUrl+"/_api/web/lists/getbytitle('User Information List')/items?$filter=Department eq '"+strdepartment+"'";


this.xhrCallforData(url);
}


  }

  private xhrCallforData(url) : any
  {
    $("#row-container").empty();

    let userArr= Array< UserObjectType>();
    $.ajax({
    url: url,
    method: "GET",
    headers: { "Accept": "application/json; odata=verbose","Content-Type": "application/json"},
               /* "X-RequestDigest": getRequestVal(webUrl)}, */
    success: function (result) {
        $("#row-container").empty();

        let listDocument=result.d.results;
       console.log(result.d.results);
        $.each(listDocument,function(index,item){
            let usrObj:UserObjectType=new UserObjectType();
            usrObj.DelveImageUrl="https://eur.delve.office.com/mt/v3/people/profileimage?userId="+encodeURIComponent(item.UserName)+"&size=L";
            usrObj.ImageUrl=item.Picture.Url;
            usrObj.FullName=item.FirstName+ " "+ item.LastName;
            usrObj.Name=item.Title;
            usrObj.Department=(item.Department!=null)?item.Department:"";
            usrObj.Email=item.EMail;
            usrObj.JobTitle=(item.JobTitle!=null)?item.JobTitle:"";
            usrObj.office=(item.Office!=null)?item.Office:"";
            usrObj.UserID=item.Name;
            usrObj.Workphone=(item.WorkPhone!=null)?item.WorkPhone:"";
            usrObj.ProfileUrl=
            "https://sandefjord365-my.sharepoint.com/PersonImmersive.aspx?accountname="+ encodeURIComponent(usrObj.UserID).replace(/%20/g,'+');
            usrObj.UserName=item.UserName;

            userArr.push(usrObj);

            let $html=`<div class="col-xs-12 col-sm-6 col-md-4">
            <div class="image-flip" ontouchstart="this.classList.toggle('hover');">
                <div class="mainflip">
                    <div class="frontside">
                        <div class="card">
                       
                            <div class="card-body text-center">`;
            let $pic= `<p><img class=" img-fluid" src="`+usrObj.DelveImageUrl+`" alt="card image"></p>`;  
            let $FullName=`<a href="`+usrObj.ProfileUrl+`" target="_blank" ><h6 class="card-title" >` +usrObj.FullName+`</h6></a>`;
            let  $email=` <p class="card-text"><i class="fa fa-envelope"></i><span>`+usrObj.Email+`</span></p>`;
            let $BackStart=`<p class="front"> <a href="#" class="flipme"><i class="fa fa-arrow-circle-right fa-2x"></i> </a></p></div>
            </div>
        </div>
        <div class="backside">
        
            <div class="card" data-href="`+usrObj.ProfileUrl+`">
            
                <div class="card-body text-center  mt-4">`;
                let $Title=`<ul class="list-inline">
                <li class="list-inline-item">
                    <a class="social-icon text-xs-center" target="_blank" href="`+usrObj.ProfileUrl+`">
                        <i class="fa fa-windows" ></i>
                    </a>
                </li>
                </ul><h5 class="card-title">`+usrObj.FullName+`</h5>`   ; 
                let $props=`<div class="well text-left">
                <ul class="list-group">
                <li class="list-group-item" data-toggle="tooltip" title="office"><i class="fa fa-building"></i>`+usrObj.office+`</li>
                <li class="list-group-item" data-toggle="tooltip" title="Job Title"><i class="fa fa-address-card"></i>`+usrObj.JobTitle+`</li>
                <li class="list-group-item"  data-toggle="tooltip" title="department"><i class="fa fa-users"></i>`+usrObj.Department+`</li>` ; 
                let $phone=(usrObj.Workphone.length>0)?` <li class="list-group-item"  data-toggle="tooltip" title="Phone"><i class="fa fa-phone-square "></i>`+usrObj.Workphone+`</li></ul></div>`:`</ul></div>`;
                let delveLink=`<p class="back"><a href="#" class="flipme"><i class="fa fa-arrow-circle-left fa-2x"> </i></a></p>
                </div>
            </div> 
             </a>
        </div>
    </div>
</div>`;     
let $usrHtml=$html+$pic+$FullName+$email+ $BackStart+$Title+$props+$phone+delveLink;
$("#row-container").append($usrHtml)  ;

        });

  $('.backside .card','.card-title').on('click',function( e){
    var profileurl=$(this).data('href');
    console.log(url);
    var win = window.open(profileurl, '_blank');
    win.focus();
  });  

  $('.flipme').on('click',function(e){
  e.preventDefault();
  $(this).parents(".mainflip").toggleClass("fliped");
  });  


    },
    error: function (data) {
       console.log(data);
    }

    
});


  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }
 
  

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void
  {
    $("#row-container").empty();

    this.context.propertyPane.refresh();

  }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    return {
      pages: [
        {
          header: {
            description: ""
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                
              PropertyPaneTextField('Department', {
                label: "Departments",
                value:''
              })
              ]
            }
          ]
        }
      ],
      loadingIndicatorDelayTime: 1,
        showLoadingIndicator: false
    };
  }

  private loadDepartment(termSets: SP.Taxonomy.TermSetCollection,spContext:SP.ClientContext )
  {

    let termSet = termSets.getByName("Department");
    let terms = termSet.getAllTerms();
    spContext.load(terms);
    spContext.executeQueryAsync(function () {
      var termsEnum = terms.getEnumerator();
      let termDepartment:any[]=[];
      while (termsEnum.moveNext()) {
        var spTerm = termsEnum.get_current();
        termDepartment.push({label:spTerm.get_name(),value:spTerm.get_name(), id:spTerm.get_id()});
      }

      window['termDepartment']= termDepartment;
      this._listOptions = [];
      let listValues=window['termDepartment'];
      var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
      if(listValues.length > 0)
      {

        options.push( { key: '0', text: "Select Department" });
             $.each(listValues,function(index,item){
              if(item.value != undefined)
              {

               options.push( { key: item.value, text: item.label });

              }
          });
      }
        this._listOptions=options;
        console.log(this._listOptions);

      });


     return  this._listOptions;
  }
}


