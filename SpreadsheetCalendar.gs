function onOpen() {
  var ui = SpreadsheetApp.getUi(); 
  // Or DocumentApp or FormApp.
  ui.createMenu('Calendar')
      .addItem('Process Calendar', 'proCalendar')
//      .addItem('Changes toggle', 'changeTrack')
      .addToUi();
//  var scriptProperties = PropertiesService.getScriptProperties();
// scriptProperties.setProperty('trackChange', 'true');
}

function onEdit(event)
{
  var scriptProperties = PropertiesService.getScriptProperties();
  var value = scriptProperties.getProperty('trackChange');
  Logger.log(value);
  if(value =='true'){
    
    var ss = event.range.getSheet();
    var changedCell = event.source.getActiveRange();
    var background = 'red';
    changedCell.setBackground(background);
  }
}


function changeTrack() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var value = scriptProperties.getProperty('trackChange');
  if(value =='true'){
 scriptProperties.setProperty('trackChange', 'false');
  } else {
 scriptProperties.setProperty('trackChange', 'true');
  }
}

function proCalendar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName('Corporate_Calendar');
  var lastRow = sheet.getLastRow();
  var j = 2 ;
  var dataRange = sheet.getRange(j, 1,lastRow , 8); //
  var data = dataRange.getValues();
  var cal = CalendarApp.getCalendarById('yourcalendarID@group.calendar.google.com'); //Calendar ID de Board Calendar
  for (i in data) {
    // Logger.log(data[i])
    
    var row = data[i]; // se lee el row con los datos del evento. 
    
    var title = row[0];  // Event Title 
    var loc = row[1];       // Event Location 
    var startDate = new Date(row[2]); // Start Date 
    var endDate = new Date(row[3]); // End Date 
    var desc = row[4]; // Description 
    var gue = row[5]; //Guest 
    var id = row[7]; // El id el evento 
    
    switch(row[6]){ // es importante para evaluar accion 

    case "Create":    
       
        //var event = cal.getEventSeriesById('eedugs8a5c9ukj71h8983d0tlk@google.com')
        //event.deleteEventSeries()
           
        var event = cal.createEvent(title, startDate, endDate , {description:desc,location:loc,guests:gue});
        var linkcell = event.getId();
        sheet.getRange(j,7).setValue("Done")
        sheet.getRange(j,8).setValue(linkcell)
        j++;
    break;
        
     case "Delete":
        cal.getEventSeriesById(id).deleteEventSeries();
        sheet.getRange(j,7).setValue("Done");
        sheet.getRange(j,8).setValue("");
        j++;
     break;
     
      case "Location":
        cal.getEventSeriesById(id).setLocation(loc);
        sheet.getRange(j,7).setValue("Done");
        j++;
      break;
      
      case "Title":
        cal.getEventSeriesById(id).setTitle(title);
        sheet.getRange(j,7).setValue("Done");
        j++;
      break;  
      
      case "Description":
        cal.getEventSeriesById(id).setDescription(desc);
        sheet.getRange(j,7).setValue("Done");
        j++;
      break;
        
      case "AddGuest":
        cal.getEventSeriesById(id).addGuest(gue);
        sheet.getRange(j,7).setValue("Done");
        j++;
      break;
      
      default:
        j++;
      break;
    
 }
   
 }
}

