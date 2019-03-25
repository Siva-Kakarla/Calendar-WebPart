import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape, trimEnd} from '@microsoft/sp-lodash-subset';

import styles from './CalendarViewWebPart.module.scss';
import * as strings from 'CalendarViewWebPartStrings';

import * as spns from "sp-pnp-js";

export interface ICalendarViewWebPartProps {
  description: string;
}

export default class CalendarViewWebPart extends BaseClientSideWebPart<ICalendarViewWebPartProps> {

  private _HTML_Output : string = "";
  private Calendar_Next_Month = new Date();
  private Calendar_Previous_Month = new Date();
  private Previus_Btn = (Math.random()).toString();
  private Next_Btn = (Math.random()).toString();

  private _MonthlyEvents:any[] ;
  
  public render(): void {
    this.Create_Month_View(null);
  }

  protected Create_Month_View(RawDate)
  {
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

    var Month_With_all_Dates = [];
    var Previous_Moth_Days = [];

    for(var i =1; i<= Index_Of_Day; i++)
    {
        var tempDay_S = this.CreateRequireDate(Month_StartDay, false, i);
        Previous_Moth_Days.push(tempDay_S);
    }
    
    Previous_Moth_Days.reverse();

    for(var i=0; i<=Previous_Moth_Days.length-1; i++)
    {
        Month_With_all_Dates.push(Previous_Moth_Days[i]);
    }
    
    Month_With_all_Dates.push(Month_StartDay);

    for(var i=1; i<=32; i++)
    {
      var tempDay_P = this.CreateRequireDate(Month_StartDay, true, i);
      Month_With_all_Dates.push(tempDay_P);

      if(Month_EndDay.getMonth() == tempDay_P.getMonth() && 
      Month_EndDay.getDate() == tempDay_P.getDate() && 
      Month_EndDay.getFullYear() == tempDay_P.getFullYear())
      {
        break;
      }
    }

    if((Month_With_all_Dates.length-1 != 41) && (Month_With_all_Dates.length-1 < 41))
    {
      var c = Month_With_all_Dates.length-1;
      do 
      {
        var tempDay1 = this.CreateRequireDate(Month_With_all_Dates[c], true, 1);
        Month_With_all_Dates.push(tempDay1);
        c++;
      }
      while (c < 41);
    }

    this._HTML_Output ="";    
    
    console.log("Get Items for Month.... ====>");

    this.apiCall(Month_With_all_Dates[0], Month_With_all_Dates[Month_With_all_Dates.length-1]).then((result)=>
    {

      this._MonthlyEvents= result;
      
      
      for(var j=0; j<=41;)
      {
        this._HTML_Output = this._HTML_Output + this.Create_Week_View(j,Month_With_all_Dates);
        j= j+7;
      }    


      const Final_HTML = `
    
        <div class="${ styles.MyCalendar}">
          
          <div class="${styles.CalendarHeadder}">        
            <div class="${styles.previous}">
              <a id="${this.Previus_Btn}" class="${styles.button}">
                <i class="ms-Icon ms-Icon--DoubleChevronLeft12" aria-hidden="true"></i>
              </a>
            </div>
            <div>
            <h2>`+ current_MonthName + `(`+ Today_year +`)`+`</h2>
            </div>
            <div class="${styles.next}">
              <a id="${this.Next_Btn}" class="${styles.button}">
                <i class="ms-Icon ms-Icon--DoubleChevronRight12" aria-hidden="true"></i>
              </a>
            </div>
          </div>

          <div id="Error_Str" style="display:block"></div>
          
          <div class='${styles.CalendarGrid}'>
            <div class="${styles.weekHead}">
              <div class="${styles.weekDay}">Sunday</div>
              <div class="${styles.weekDay}">Monday</div>
              <div class="${styles.weekDay}">Tuesday</div>
              <div class="${styles.weekDay}">Wednesday</div>
              <div class="${styles.weekDay}">Thursday</div>
              <div class="${styles.weekDay}">Friday</div>
              <div class="${styles.weekDay}">Saturday</div>  
            </div>

            `+ this._HTML_Output +`

          </div>`;

      console.log("<=======   Form Final HTML>");
      
      this.domElement.innerHTML = Final_HTML;

      document.getElementById(this.Previus_Btn).addEventListener("click", this.Previous_ClickHandler.bind(this));
      document.getElementById(this.Next_Btn).addEventListener("click", this.Next_ClickHandler.bind(this));

    });
  }

  private Create_Week_View(x, FinalArray)
  {
    var _Weekly_Items = this.Get_Week_Events(x, FinalArray);
    console.log(_Weekly_Items);
    var Week_Event_Html = this.Create_Week_Events(_Weekly_Items);
    
    return  `
      <div class="${styles.week}">
        <div class="${styles.day}">
          <h3 class="${styles.dayLabel}">`+ FinalArray[x].getDate() +`</h3>
        </div>
        <div class="${styles.day}">
          <h3 class="${styles.dayLabel}">`+ FinalArray[(1+ x)].getDate() +`</h3>
        </div>
        <div class="${styles.day}">
          <h3 class="${styles.dayLabel}">`+ FinalArray[(2+ x)].getDate() +`</h3>
        </div>
        <div class="${styles.day}">
          <h3 class="${styles.dayLabel}">`+ FinalArray[(3+ x)].getDate() +`</h3>
        </div>
        <div class="${styles.day}">
          <h3 class="${styles.dayLabel}">`+ FinalArray[(4+ x)].getDate() +`</h3>
        </div>
        <div class="${styles.day}">
          <h3 class="${styles.dayLabel}">`+ FinalArray[(5+ x)].getDate() +`</h3>
        </div>
        <div class="${styles.day}">
          <h3 class="${styles.dayLabel}">`+ FinalArray[(6+ x)].getDate() +`</h3>
        </div>
        ${ Week_Event_Html }
      </div>`;
  }

  private Create_Week_Events(_Weekly_Items)
  {
    var _Event_Of_Day_HTML = "";

    _Weekly_Items.forEach(_Week_Item => 
      {
        debugger;
        if(_Week_Item.Start_InWeek == true && _Week_Item.End_InWeek == true)
        {
          _Event_Of_Day_HTML += `<div class="${styles.eventStartEnd}" data-span="`+_Week_Item.Days.length+`" style="grid-column-start: `+_Week_Item.Days[0] +`; grid-column-end: span `+_Week_Item.Days.length+`; height:15px; font-size:x-small;">`+_Week_Item.Title+`</div>`;   
        }
        else if(_Week_Item.Start_InWeek == true && _Week_Item.End_InWeek == false)
        {
          _Event_Of_Day_HTML += `<div class="${styles.eventStart}" data-span="`+_Week_Item.Days.length+`" style="grid-column-start: `+_Week_Item.Days[0] +`; grid-column-end: span `+_Week_Item.Days.length+`; height:15px; font-size:x-small;">`+_Week_Item.Title+`</div>`;
        }
        else if(_Week_Item.Start_InWeek == false && _Week_Item.End_InWeek == true)
        {
          _Event_Of_Day_HTML += `<div class="${styles.eventEnd}" data-span="`+_Week_Item.Days.length+`" style="grid-column-start: `+_Week_Item.Days[0] +`; grid-column-end: span `+_Week_Item.Days.length+`; height:15px; font-size:x-small;">`+_Week_Item.Title+`</div>`;
        }
        else if(_Week_Item.Start_InWeek == false && _Week_Item.End_InWeek == false)
        {
          _Event_Of_Day_HTML += `<div class="${styles.event}" data-span="`+_Week_Item.Days.length+`" style="grid-column-start: `+_Week_Item.Days[0] +`; grid-column-end: span `+_Week_Item.Days.length+`; height:15px; font-size:x-small;">`+_Week_Item.Title+`</div>`;
        }
      });

      // <div class="${styles.event}" data-span="1" style="grid-column-start: 4; grid-column-end: span 1; height:15px; font-size:x-small;">+ More</div>

    return _Event_Of_Day_HTML;
  }

  private Get_Week_Events(x, FinalArray)
  {
    var _Weekly_Items = [];
    var FirstDay_Of_Week = FinalArray[x];
    var LastDay_Of_Week = FinalArray[x+6];
    
    for(var i_day=1; i_day<=7; i_day++)
    {
      this._MonthlyEvents.forEach(Current_Day_Event => 
      {
          var StartDay:any = new Date(Current_Day_Event["StartDay"]);
          var EndDay:any = new Date(Current_Day_Event["EndDay"]);
          var currentDay:any = FinalArray[x];
          let newName;
          var IsStar_In_Week = true;
          var IsEnd_In_Week = true;
          
          if((StartDay.getTime() <= currentDay.getTime()) && (EndDay.getTime() >= currentDay.getTime()))
          {
            if((StartDay.getTime() >= FirstDay_Of_Week.getTime()) && (StartDay.getTime() <= LastDay_Of_Week.getTime()))
            {
              IsStar_In_Week = true;
            }
            else
            {
              IsStar_In_Week = false;
            }
            if(EndDay.getTime() <= LastDay_Of_Week.getTime())
            {
              IsEnd_In_Week = true;
            }
            else
            {
              IsEnd_In_Week = false;
            }
            newName = {
                Id:Current_Day_Event["Id"],
                Title:Current_Day_Event["Title"],
                Days:i_day,
                Start_InWeek: IsStar_In_Week,
                End_InWeek: IsEnd_In_Week
              };
              _Weekly_Items.push(newName);
          }
      });
      x++;      
    }

    //Get Unique values from Raw Array.
    _Weekly_Items = _Weekly_Items.filter((value, index, array) => 
     !array.filter((v, i) => JSON.stringify(value) == JSON.stringify(v) && i < index).length);

    var Day_Chain = [];
    var _FinalArray = [];

    for(var i=0; i<_Weekly_Items.length; i++)
    {
      Day_Chain = [];
      for(var j=0; j<_Weekly_Items.length; j++)
      {
        if((_Weekly_Items[i].Id) === (_Weekly_Items[j].Id))
        {
          Day_Chain.push(_Weekly_Items[j].Days);
        }
      }

      var newName = {
          Id:_Weekly_Items[i].Id,
          Title:_Weekly_Items[i].Title,
          Days:Day_Chain,
          Start_InWeek: _Weekly_Items[i].Start_InWeek,
          End_InWeek:_Weekly_Items[i].End_InWeek
        };
        _FinalArray.push(newName);
    }

    _Weekly_Items = _FinalArray.filter((value, index, array) => 
     !array.filter((v, i) => JSON.stringify(value) == JSON.stringify(v) && i < index).length);

     return _Weekly_Items;
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

  private apiCall(StartDay, EndDay):Promise<any>{
    return new Promise(resolve => 
    {

      var month = StartDay.getMonth();//d.getMonth()+1; (For Get month number 1 to 12 insted 0 to 11).
      var day = StartDay.getDate();
      var year = StartDay.getFullYear();

      var Start_Date = year + '-' + ((month+1)<10 ? '0' : '') + (month+1) + '-' + (day<10 ? '0' : '') + 
      day + 'T00:00:00Z';

      month = EndDay.getMonth();//d.getMonth()+1; (For Get month number 1 to 12 insted 0 to 11).
      day = EndDay.getDate();
      year = EndDay.getFullYear();

      var End_Date = year + '-' + ((month+1)<10 ? '0' : '') + (month+1) + '-' + (day<10 ? '0' : '') + day + 'T00:00:00Z';

      const xml = `    
      <View><ViewFields><FieldRef Name='ID' /><FieldRef Name='Title' /><FieldRef Name='StartDay' /><FieldRef Name='EndDay' /></ViewFields>
        <Query>

          <Where>
              <And>
                <Or>
                    <Geq>
                      <FieldRef Name='StartDay' />
                      <Value IncludeTimeValue='TRUE' Type='DateTime'>`+ Start_Date +`</Value>
                    </Geq>
                    <Leq>
                      <FieldRef Name='StartDay' />
                      <Value IncludeTimeValue='TRUE' Type='DateTime'>`+ End_Date +`</Value>
                    </Leq>
                </Or>
                <Or>
                    <Geq>
                      <FieldRef Name='EndDay' />
                      <Value IncludeTimeValue='TRUE' Type='DateTime'>`+ Start_Date +`</Value>
                    </Geq>
                    <Leq>
                      <FieldRef Name='EndDay' />
                      <Value IncludeTimeValue='TRUE' Type='DateTime'>`+ End_Date +`</Value>
                    </Leq>
                </Or>
              </And>
          </Where>


          <OrderBy>
              <FieldRef Name='StartDay' Ascending='False' />
          </OrderBy>
        </Query>
      </View>
      `;

      const q: spns.CamlQuery = {
          ViewXml: xml,
      };

      spns.sp.web.lists.getByTitle("TestListCalendar").getItemsByCAMLQuery(q).then((r: any[]) => {
        //console.log(JSON.stringify(r, null, 4));
        resolve(r);

      }).catch(function(err)
      {
        document.getElementById("Error_Str").innerText=err;
        resolve(null);
      });
    });
  }

  protected Previous_ClickHandler()
  {
    this.Create_Month_View(this.Calendar_Previous_Month);
    return false;
  }

  protected Next_ClickHandler()
  {
    this.Create_Month_View(this.Calendar_Next_Month);
    return false;
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