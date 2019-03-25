import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MarsCalendarWebPart.module.scss';

import * as strings from 'MarsCalendarWebPartStrings';

import * as $ from 'jquery';

import * as spns from "sp-pnp-js";

export interface IMarsCalendarWebPartProps {
  description: string;
}

export default class MarsCalendarWebPart extends BaseClientSideWebPart<IMarsCalendarWebPartProps> {

  private _HTML_Output : string = "";
  private Calendar_Next_Month = new Date();
  private Calendar_Previous_Month = new Date();

  public render(): void {

     this.ShowCalendar(null);


    //  `
    //   <div class="${ styles.marsCalendar }">
    //     <div class="${ styles.container }">
    //       <div class="${ styles.row }">
    //         <div class="${ styles.column }">
    //           <span class="${ styles.title }">Welcome to SharePoint!</span>
    //           <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
    //           <p class="${ styles.description }">${escape(this.properties.description)}</p>
    //           <a href="https://aka.ms/spfx" class="${ styles.button }">
    //             <span class="${ styles.label }">Learn more</span>
    //           </a>
    //         </div>
    //       </div>
    //     </div>
    //   </div>`;
  }

  private ShowCalendar(O_Date)
  {
    this.domElement.innerHTML = this.CreateCalendarView(O_Date);

    document.getElementById('Previous_Btn').addEventListener("click", this.Previous_onClickHandler.bind(this));
    document.getElementById('Next_Btn').addEventListener("click", this.Next_onClickHandler.bind(this)); 
  }

  protected CreateCalendarView(RawDate)
  {
    debugger;
    var weekday = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
    
    var monthNames = ["January", "February", "March", 
    "April", "May", "June", "July", "August", 
    "September", "October", "November", "December"];

    var Base_Date = null;

    if(RawDate != null)
    {
      Base_Date = RawDate;
    }
    else
    {
      Base_Date = new Date();
    }

    var month = Base_Date.getMonth();//d.getMonth()+1; (For Get month number 1 to 12 insted 0 to 11).
    var day = Base_Date.getDate();
    var year = Base_Date.getFullYear();

    var current_MonthName = monthNames[month];

    var Today_Date = year + '/' + ((month+1)<10 ? '0' : '') + (month+1) + '/' + (day<10 ? '0' : '') + day;
    var Today_year = year;

    var Month_StartDay = new Date(year, month, 1);
    var Month_EndDay = new Date(year, month+1, 0);
    var StartDay_Week = weekday[Month_StartDay.getDay()];
    var Index_Of_Day = weekday.indexOf(StartDay_Week);

    var tempDay_L = this.CreateRequireDate(Month_StartDay, false, 1);
    this.Calendar_Previous_Month = new Date(tempDay_L.getFullYear(), tempDay_L.getMonth(), 1);

    var tempDay_G = this.CreateRequireDate(Month_EndDay, true, 1);
    this.Calendar_Next_Month = new Date(tempDay_G.getFullYear(), tempDay_G.getMonth(), 1);

    var FinalDaysinCurrentMonth = [];
    var Previous_Moth_Days = [];

    for(var i =1; i<= Index_Of_Day; i++)
    {
        var tempDay_S = this.CreateRequireDate(Month_StartDay, false, i);
        Previous_Moth_Days.push(tempDay_S);
    }
    
    Previous_Moth_Days.reverse();

    for(var i=0; i<=Previous_Moth_Days.length-1; i++)
    {
        FinalDaysinCurrentMonth.push(Previous_Moth_Days[i]);
    }
    
    FinalDaysinCurrentMonth.push(Month_StartDay);

    for(var i=1; i<=32; i++)
    {
      var tempDay_P = this.CreateRequireDate(Month_StartDay, true, i);
      FinalDaysinCurrentMonth.push(tempDay_P);

      if(Month_EndDay.getMonth() == tempDay_P.getMonth() && 
      Month_EndDay.getDate() == tempDay_P.getDate() && 
      Month_EndDay.getFullYear() == tempDay_P.getFullYear())
      {
        break;
      }
    }

    if((FinalDaysinCurrentMonth.length-1 != 41) && (FinalDaysinCurrentMonth.length-1 < 41))
    {
      var c = FinalDaysinCurrentMonth.length-1;
      do 
      {
        var tempDay1 = this.CreateRequireDate(FinalDaysinCurrentMonth[c], true, 1);
        FinalDaysinCurrentMonth.push(tempDay1);
        c++;
      }
      while (c < 41);
    }

    this._HTML_Output ="";    

    for(var j=0; j<=41;)
    {
      this._HTML_Output = this._HTML_Output + this.BuildDay(j,FinalDaysinCurrentMonth);
      j= j+7;
    }   //&laquo;       &raquo;

    const Final_HTML = `
    
    <div class="${ styles.MyCalendar}">
      
      <div class="${styles.CalendarHeadder}">        
        <div class="${styles.previous}">
          <a id="Previous_Btn" class="${styles.button}">
            <i class="ms-Icon ms-Icon--DoubleChevronLeft12" aria-hidden="true"></i>
          </a>
        </div>
        
        <div>
          <h2>`+ current_MonthName + `(`+ Today_year +`)`+`</h2>
        </div>
        
        <div class="${styles.next}">
          <a id="Next_Btn" class="${styles.button}">
            <i class="ms-Icon ms-Icon--DoubleChevronRight12" aria-hidden="true"></i>
          </a>
        </div>
      </div>

      <div class="ms-acal-error" id="WPQ2_err" style="display:none"></div>
      
      <table class="${ styles.CustomeTable }">
        <tbody>        
          <tr class="${styles.rowHead}">
            <th class="${ styles.th }">Sunday</th>
            <th class="${ styles.th }">Monday</th>
            <th class="${ styles.th }">Tuesday</th>
            <th class="${ styles.th }">Wednesday</th>
            <th class="${ styles.th }">Thursday</th>
            <th class="${ styles.th }">Friday</th>
            <th class="${ styles.th }">Saturday</th>
          </tr>

          `+ this._HTML_Output +`
        </tbody>
      </table>
    </div>`;

    return Final_HTML;
  }//
  //<caption>${Today_Date}</caption>
  // <p>`+ Today_Date +`</p>

  private BuildDay(x, FinalArray)
  {
    var day1 = x;
    x++;
    var day2 = x;
    x++;
    var day3 = x;
    x++;
    var day4 = x;
    x++;
    var day5 = x;
    x++;
    var day6 = x;
    x++;
    var day7 = x;
      return  `     
      <tr class="${styles.rowData}">
        <td class="${ styles.td }">`+ FinalArray[day1].getDate() +`</td>
        <td class="${ styles.td }">`+ FinalArray[day2].getDate() +`</td>
        <td class="${ styles.td }">`+ FinalArray[day3].getDate() +`</td>
        <td class="${ styles.td }">`+ FinalArray[day4].getDate() +`</td>
        <td class="${ styles.td }">`+ FinalArray[day5].getDate() +`</td>
        <td class="${ styles.td }">`+ FinalArray[day6].getDate() +`</td>
        <td class="${ styles.td }">`+ FinalArray[day7].getDate() +`</td>			
      </tr>`;
  }

  private CreateRequireDate(InPutDate, IsAddDys, NumberOfDays)
  {
    try
    {
      var month_of_Date = InPutDate.getMonth();
      var day_of_Date = InPutDate.getDate();
      var year_of_Date = InPutDate.getFullYear();
      var Temp_Date_String = (month_of_Date + 1) + "/" + day_of_Date + "/" + year_of_Date;
      var Final_Date = new Date(Temp_Date_String);

      if(IsAddDys == true)
      {
        Final_Date.setDate(Final_Date.getDate() + NumberOfDays);
      }
      else
      {
        Final_Date.setDate(Final_Date.getDate() - NumberOfDays);
      }

      return Final_Date;
    }
    catch(ex)
    {
      return null;
    }
  }

  protected Previous_onClickHandler()
  {
    debugger;
    this.ShowCalendar(this.Calendar_Previous_Month);
    return false;
  }

  protected Next_onClickHandler()
  {
    this.ShowCalendar(this.Calendar_Next_Month);
    return false;
  }

  private GetListItem(URL)
  {
    let w = new spns.Web(URL);//"{Absolute Web Path}"
    w.get().then(w => { });

    spns.sp.web.lists.getByTitle("Tasks").items.get()
    .then(function(data)
    {
      document.getElementById("main").innerText=data.Title;
    })
    .catch(function(err)
    {
      document.getElementById("main").innerText=err;
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
