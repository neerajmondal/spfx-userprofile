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
require("popper.js");
require("bootstrap");
import * as popper from 'popper.js';
import * as bootstrap from 'bootstrap';

require('mdbootstrap');
require('./basestyles.css');

const imgdefaultUser: any = require('./assets/defaultuser.png');
const imgDelveIcon: any = require('./assets/delve.png');

let userArr= Array< UserObjectType>();

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

  public usrObj:UserObjectType

  constructor()
  {
    super();

    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css');
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css');
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/mdbootstrap/4.5.4/css/mdb.min.css');
  


  }
  protected onInit(): Promise<void> {


    return super.onInit();
  }



  public render(): void {

    if (!this.renderedOnce) {
    this.domElement.innerHTML = `
    <section class="team pb-5">
          <div class="container">
             <!-- <h4 class="section-title"> Department</h4>-->
             
                 <div id="row-container" class="row">


                  </div>

                  <!-- Full Height Modal Right Success Demo-->
<div  class="modal fade right" id="fluidModalRightSuccessDemo" tabindex="-1" role="dialog" aria-labelledby="myModalLabel"
    aria-hidden="true" data-backdrop="false">
    <div class="modal-dialog modal-full-height modal-right modal-notify modal-info" role="document">
        <!--Content-->
        <div class="modal-content">
            <!--Header-->
            <div class="modal-header">
                <p class="heading lead"></p>

                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                    <span aria-hidden="true" class="white-text">&times;</span>
                </button>
            </div>

            <!--Body-->
            <div class="modal-body nopadding">
                <div class="text-center">
                    
                   
                </div>
               
            </div>

            <!--Footer-->
            <div class="modal-footer justify-content-center">
            <button type="button" class="btn btn-info btn-sm" data-dismiss="modal">Close</button>
            </div>
        </div>
        <!--/.Content-->
    </div>
</div>
<!-- Full Height Modal Right Success Demo-->
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

  
    userArr.splice(0, userArr.length);

    $.ajax({
    url: url,
    method: "GET",
    headers: { "Accept": "application/json; odata=verbose","Content-Type": "application/json"},
               /* "X-RequestDigest": getRequestVal(webUrl)}, */
    success: function (result) {
        $("#row-container").empty();

        let listDocument=result.d.results;
       //console.log(result.d.results);
        $.each(listDocument,function(index,item){
            let usrObj=new UserObjectType();
            usrObj.DelveImageUrl="https://eur.delve.office.com/mt/v3/people/profileimage?userId="+encodeURIComponent(item.UserName)+"&size=L";
            if(item.Picture!=null){
            usrObj.ImageUrl=item.Picture.Url;
            if(usrObj.ImageUrl.indexOf("MThumb") >=0){
            usrObj.ImageUrl =usrObj.ImageUrl.replace("MThumb","LThumb");
            }
          }
          else{
            usrObj.ImageUrl=imgdefaultUser;
          }
            usrObj.FullName=item.FirstName+ " "+ item.LastName;
            usrObj.Name=item.Title;
            usrObj.Department=(item.Department!=null)?item.Department:"";
            usrObj.Email=item.EMail;
            usrObj.JobTitle=(item.JobTitle!=null)?item.JobTitle:"";
            usrObj.office=(item.Office!=null)?item.Office:"";
            usrObj.UserID=item.Name;
            usrObj.Workphone=(item.WorkPhone!=null)?item.WorkPhone:"";
            usrObj.ProfileUrl={LinkUrl:
            "https://sandefjord365-my.sharepoint.com/PersonImmersive.aspx?accountname="+ encodeURIComponent(usrObj.UserID).replace(/%20/g,'+'),
        DetailsUrl:"/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='"+encodeURIComponent(usrObj.UserID).replace(/%20/g,'+')+"'"
        }
            usrObj.UserName=item.UserName;

            userArr.push(usrObj);

            let $html=`<div class="col-xs-12 col-sm-6 col-md-4">
            <div class="image-flip" ontouchstart="this.classList.toggle('hover');">
                <div class="mainflip">
                    <div class="frontside">
                        <div class="card" data-index=${index} data-detailurl="${ usrObj.ProfileUrl.DetailsUrl}">
                       
                            <div class="card-body text-center">`;
            let $pic= `<p><img class=" img-fluid" src="${usrObj.ImageUrl}" alt="card image"></p>`;  
            let $FullName=`<a href="${usrObj.ProfileUrl.LinkUrl}" target="_blank" ><h6 class="card-title" >${usrObj.FullName}</h6></a>`;
            let  $email=` <div class="card-text"><i class="fa fa-envelope"></i><span>${usrObj.Email}</span></div><div class="front"> <a href="#" class="flipme"><i class="fa fa-arrow-circle-right fa-2x"></i> </a></div>
            </div>
           
            </div>
        </div>`;
            let $BackStart=` <div class="backside">
            <div class="card">
            
                <div class="card-body text-center  mt-4">`;
                let $Title=` <div class='${styles["card-links"]}'>
                    <a class="social-icon text-xs-center" target="_blank" href="${usrObj.ProfileUrl.LinkUrl}">
                        <i class="fa fa-windows" ></i>
                    </a>
               </div><h5 class="card-title">${usrObj.FullName}</h5>`   ; 
                let $props=`<div class="well text-left">
                <ul class="event-list">
                     <li  data-toggle="tooltip" data-animation="false" title="office" >
                         
                         <i class="fa fa-building  icon" > </i>
                         <div class="info">
                             <p class="desc">${usrObj.office}</p>
                         </div>
                         
                     </li>
                     <li  data-toggle="tooltip" data-animation="false" title="Job Title" >
                         
                     <i class="fa fa-id-badge  icon" > </i>
                     <div class="info">
                         <p class="desc">${usrObj.JobTitle}</p>
                     </div>
                     
                 </li>
                 </li>
                 <li  data-toggle="tooltip" data-animation="false" title="Department" >
                     
                 <i class="fa fa-users icon" > </i>
                 <div class="info">
                     <p class="desc">${usrObj.Department}</p>
                 </div>
                 
             </li></ul></div><p class="back"><a href="#" class="flipme"><i class="fa fa-arrow-circle-left fa-2x"> </i></a></p>` ; 
               
                let delveLink=`</div>
            </div> 
             </a>
        </div>
    </div>
</div>`;     
let $usrHtml=$html+$pic+$FullName+$email+ $BackStart+$Title+$props+delveLink;
$("#row-container").append($usrHtml)  ;

 });

    $('img.img-fluid').on('error',function(){
                $(this).attr('src', imgdefaultUser);
    });          
  $('.frontside .card').on('click',function( e){
    var detailurl=$(this).data('detailurl');
    var index=$(this).data('index');
        $.ajax({
            url: detailurl,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose","Content-Type": "application/json"},
                    /* "X-RequestDigest": getRequestVal(webUrl)}, */
            success: function (result) {
                let resultJson=result.d;
                let profileCard:string,profileDetails:string;
                let  userDetails:any={};

                if(result.d.hasOwnProperty('GetPropertiesFor')){
                    userDetails.PictureUrl={Title: "Profile Picture Url" , Value:imgdefaultUser};
                    userDetails.DisplayName={ Title: "Name" , Value: userArr[index].FullName};
                    userDetails.Email={Title: "Email" ,  Value: userArr[index].Email};
                }
               else{

                userDetails.DisplayName={ Title: "Name" , Value:resultJson.DisplayName};
                userDetails.Email={Title: "Email" , Value:resultJson.Email};


                userDetails.WorkPhone={Title: "Work Phone" , Value:resultJson.UserProfileProperties.results.find(x => x.Key === 'WorkPhone').Value};
                userDetails.Department={Title: "Department" , Value:resultJson.UserProfileProperties.results.find(x => x.Key === 'Department').Value};
                userDetails.AboutMe={Title: "About Me" , Value:resultJson.UserProfileProperties.results.find(x => x.Key === 'AboutMe').Value};
                userDetails.JobTitle={Title: "Job Title" , Value:resultJson.UserProfileProperties.results.find(x => x.Key === 'SPS-JobTitle').Value}
                userDetails.PictureUrl={Title: "Profile Picture Url" , Value:resultJson.UserProfileProperties.results.find(x => x.Key === 'PictureURL').Value};
                userDetails.Office={Title: "Office" , Value:resultJson.UserProfileProperties.results.find(x => x.Key === 'Office').Value};
                userDetails.Location={Title: "Location" , Value:resultJson.UserProfileProperties.results.find(x => x.Key === 'SPS-Location').Value};
                userDetails.Skils={Title: "Skills" , Value:resultJson.UserProfileProperties.results.find(x => x.Key === 'SPS-Skills').Value};
                userDetails.Interest={Title: "Interests" , Value:resultJson.UserProfileProperties.results.find(x => x.Key === 'SPS-Interests').Value};
                userDetails.MsOnlineId={Title: "Msonline" , Value:resultJson.UserProfileProperties.results.find(x => x.Key === 'msOnline-ObjectId').Value};

                userDetails.DelveProfile="https://eur.delve.office.com/?u="+userDetails.MsOnlineId.Value+"&v=profiledetails";

                if(userDetails.PictureUrl.Value.indexOf("MThumb") >=0){
                    userDetails.PictureUrl.Value =userDetails.PictureUrl.Value.replace("MThumb","LThumb");
                    }
               }

               

                 profileCard=`<div class='card ${styles["card-profile"]} text-center'>
                <img alt='' class='${styles["card-img-top"]}' src='/SiteAssets/Hval_2.jpg'>
                <div class='card-block'>
                  <img alt='' class='${styles["card-img-profile"]}' src='${userDetails.PictureUrl.Value}'>
                  <h4 class='${styles["card-title"]}'>
                    ${userArr[index].Name}
                   ` ;


                   if(result.d.hasOwnProperty('GetPropertiesFor')){
                    profileDetails=` 
                    <small><i class="fa fa-envelope"></i>${userArr[index].Email}</small>
                  </h4><ul class="event-list">
					<li>
						
						<i class="fa fa-wpexplorer icon" > </i>
						<div class="info">
							<h4 class="title">Information not Available</h4>
							<p class="desc"></p>
						</div>
						
                    </li></ul> </div>
                    </div>`;
                  
                  }
                  else{

                    profileDetails =`<small>${userDetails.AboutMe.Value}</small>
                    <small><i class="fa fa-envelope"></i>${userArr[index].Email}</small>
                  </h4><ul class="event-list">
                     <li>
                         
                         <i class="fa fa-building  icon" > </i>
                         <div class="info">
                             <h4 class="title">${userDetails.Office.Title}</h4>
                             <p class="desc">${userDetails.Office.Value}</p>
                         </div>
                         
                     </li>
                     <li>
                         
                         <i class="fa fa-users  icon" > </i>
                         <div class="info">
                             <h4 class="title">${userDetails.Department.Title}</h4>
                             <p class="desc">${userDetails.Department.Value}</p>
                         </div>
                         
                     </li>
                     <li>
                         
                         <i class="fa fa-id-badge  icon" > </i>
                         <div class="info">
                             <h4 class="title">${userDetails.JobTitle.Title}</h4>
                             <p class="desc">${userDetails.JobTitle.Value}</p>
                         </div>
                         
                     </li>
                     <li>
                         
                     <i class="fa fa-map-marker icon" > </i>
                     <div class="info">
                         <h4 class="title">${userDetails.Location.Title}</h4>
                         <p class="desc">${userDetails.Location.Value}</p>
                     </div>
                     
                 </li>
                 <li>
                         
                 <i class="fa fa-podcast icon" > </i>
                 <div class="info">
                     <h4 class="title">${userDetails.Interest.Title}</h4>
                     <p class="desc">${userDetails.Interest.Value}</p>
                 </div>
                 
                 </li>
                 <li>
                         
                 <i class="fa fa-phone-square  icon" > </i>
                 <div class="info">
                     <h4 class="title">${userDetails.WorkPhone.Title}</h4>
                     <p class="desc">${userDetails.WorkPhone.Value}</p>
                 </div>
                 
                 </li>
                     </ul>
                   <div class='${styles["card-links"]}'>
                     <a  href='${userDetails.DelveProfile}'  target="_blank"><img src='${imgDelveIcon}' alt='' style="width:50px" /></a>
                    
                   </div>
                 </div>
               </div>`;

                  }

              $('#fluidModalRightSuccessDemo .modal-header> .heading').html( userArr[index].FullName );
              $('#fluidModalRightSuccessDemo .modal-body > div.text-center').html(profileCard+ profileDetails);

               $('#fluidModalRightSuccessDemo').modal('show'); 

            },
        
            error: function (data) {
            console.log(data);
            }
        });
    

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

  
}


