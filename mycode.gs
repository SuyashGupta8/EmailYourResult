
 /*formSubmitReply is my main fuction where other functions are called.FormSubmitReply will trigger when submit button is clicked on my form.
Google has made inbuilt events and action listener.I have to just go on Resourses on menu bar and then to current project trigger to add listener to my form submit button.*/

  function formSubmitReply(E){
    
 
       var ssform=SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1o7qXTEI9ixK_Z-Afkb8WBGlj0r1yXEaham0rxdbzRow/edit#gid=736438352');//getting ssform as active(form info) spread sheet
       var sheet = ssform.getSheets()[0];

       var lastRow=getSheetProperty(sheet).lastRow; //getting last row(intiger or number) by my object method
       var formRange=sheet.getRange(lastRow,1,1,getSheetProperty(sheet).numColumn).getValues();
          
     
         var indexFunction=gettingIndexValue(sheet,null,"Choose your semester"); //using my closure function to get index values in spreadsheet  of (column name string, here"Choose your semester") passing argument
         var indexObj=indexFunction();
         var columnNameIndex=indexObj.columnNameIndex;//we will get column index for string "Choose your semester" from formspreadsheet
    
          indexFunction=gettingIndexValue(sheet,null,"Enter your 5 digit roll number");
          indexObj=indexFunction();
          var rollColumnIndex=indexObj.columnNameIndex; //will give column index for string "Enter your 5 digit roll number" from formspreadsheet
        
    
       var formSemesterRoll;//accesses roll number values
       formSemesterRoll=formRange[0][rollColumnIndex];     
                                                                          /*Browser.msgBox(formSemesterRoll.toSource());//to display msg remove comment// useful for debugging
                                                                          Logger.log(formSemester);//to display msg*/
       
    var formSemester=formRange[0][columnNameIndex]; //gets the semester whose result we want to display
    
    //var emailId=formRange[0][1];//can use both method
    var userEmail=E.values[1];
    
                                                                    //var userEmail="suguptayash@gmail.com";//**to verify my code uncomment it
    
    if(verifyFormEntry(formSemesterRoll,formSemester))  //calling verifyformentry (will verify data fromstudentinfo spreadsheet)
    {  
     
      if(getData(formSemesterRoll,formSemester,"Roll Number") === null)   //**if we have put wrong entry in semester spreadshet(note it will check individual semester sheet)   
      {
        MailApp.sendEmail(userEmail,"Sorry","Sorry, our data sheets are under maintanance ");
      }
      
      MailApp.sendEmail(userEmail,formSemester,"",{htmlBody: getHtmlFile(formSemesterRoll,formSemester)});
       
Logger.log(userEmail);
    }
    else
    {
      MailApp.sendEmail(userEmail,"Invalid","Please enter valid informatin.Check your roll number correctly.");
    }
    
      
}​    
  
 

function verifyFormEntry(entry1,entry2){  //entry1 checking for only roll number //entry2 for semester number from (studentinfo spreadsheet)
                             
  
                  var ss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheet/ccc?key=0Aj2GJTiRHxZPdFNFOTk5eVFEWDNuZW4zYU1Lc2ZUaFE&usp=drive_web#gid=0');//getting ss as  spread sheet for student info 
                  var inSheet =ss.getSheets()[0];    
                                    //var sheetRange=ss.getRange("A1:D8");
                                 //var sheetData=sheetRange.getValues();
 
    
      if(gettingCellValue(inSheet,entry1,entry2) === "Available" )//cross chrcking student info 1.first by cheking entry1 2.second by checking entry2 3.by checkiing cell value
        
      {
        return true;
      
    }
  
else{
      return false;
    } 
  
    
  }



function getHtmlFile(RollNo,Sem){
  
  var studentInfo=SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheet/ccc?key=0Aj2GJTiRHxZPdFNFOTk5eVFEWDNuZW4zYU1Lc2ZUaFE#gid=0');
  var sheet =  studentInfo.getSheets()[0];
  
  
   var htmlFile='<html>'
                    +'<head>'
                     +'<title></title>'
                     +'</head>'
                     +'<body>'
                     +"<p >College Name: <em>"+gettingCellValue(sheet,RollNo,"College Name")+"</em></p>"  
                     
                     +"<p>"
                         +"<A2>"+"Name:"+gettingCellValue(sheet,RollNo,"Student Name")+"</br>"+"</A2>"
                       +"</p>"
                       +"<p>"
                         +"<A2>" +"Fathers Name:"+gettingCellValue(sheet,RollNo,"Father Name")+"</A2>"
                       +"</p>"
                     +"<table color:red border=1px>"
                             +"<thead>"
                                +"<tr>"
                               
                                  +"<th colspan=3>"+Sem+"</th>" 
                                +"</tr>" 
                              
                                +"<tr>"
                               
                                  +"<th colspan=3>"+"Enrollment Number:"+gettingCellValue(sheet,RollNo,"Enrollment Number")+"</th>" 
                                +"</tr>"    
                                 
                                  +"<tr>"
                                     +"<td>Subjects</td>"
                                 
                                     +"<td>Marks</td>"
                                     +"<td>Max Marks</td>"
                                     
                                  +"</tr>"
                             
                                 
                             +"</thead>"
                             +"<tr>"
                             +"<td color:green>"+gettingHeaderValues(Sem)[0][3]+"</td>"
                                  +"<td>"+getData(RollNo,Sem,gettingHeaderValues(Sem)[0][3])+"</td>"
                                  +"<td>"+getData(RollNo,Sem,gettingHeaderValues(Sem)[0][7])+"</td>"
                                 
                                                               
                               +"</tr>"
                              +"<tr>"
                                  +"<td>"+gettingHeaderValues(Sem)[0][4]+"</td>"
                                  +"<td>"+getData(RollNo,Sem,gettingHeaderValues(Sem)[0][4])+"</td>"
                                  +"<td>"+getData(RollNo,Sem,gettingHeaderValues(Sem)[0][7])+"</td>"
                                  
                            +"</tr>"
                            +"<tr>"
                                  +"<td>"+gettingHeaderValues(Sem)[0][5]+"</td>"
                                  +"<td>"+getData(RollNo,Sem,gettingHeaderValues(Sem)[0][5])+"</td>"
                                  +"<td>"+getData(RollNo,Sem,gettingHeaderValues(Sem)[0][7])+"</td>"
                                   
                            +"</tr>"
                            +"<tr>"
                                  +"<td>"+gettingHeaderValues(Sem)[0][6]+"</td>"
                                  +"<td>"+getData(RollNo,Sem,gettingHeaderValues(Sem)[0][6])+"</td>"
                                  +"<td>"+getData(RollNo,Sem,gettingHeaderValues(Sem)[0][7])+"</td>"
                               +"</tr>"      
                                  
                               +"<tr>"
                                  +"<td>"+gettingHeaderValues(Sem)[0][8]+"</td>"
                                  +"<td>"+getData(RollNo,Sem,gettingHeaderValues(Sem)[0][8])+"</td>"
                                  +"<td>"+getData(RollNo,Sem,gettingHeaderValues(Sem)[0][9])+"</td>"
                                  
                            +"</tr>"
                            
                        +"</table>"
                        +'<p>' +'</p>'
                         +'<p></p>'
                      +'</body>' 
                     +'</html>';
  
  
  /* Logger.log(getData(80000,"First semester result",gettingHeaderValues("First semester result")[0][1]));//to cross check code remove the comment and enter proper argument
  
 if( getData(80000,"First semester result",gettingHeaderValues("First semester result")[0][1]) === null)
  {
    Logger.log("null");
  }
   else
  {
    Logger.log("not nulll");
  }
  Logger.log(htmlFile);*/
 
  
  
  
  return htmlFile;  

  
}

function getData(rollno,semester,columnName){  
  
 var sheet;                                              //opening the respective semester sheet on semester result basis  
 var inSheet; 
 switch(semester){                                           
                        case "First semester result":
                        sheet= SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheet/ccc?key=0Aj2GJTiRHxZPdHRmbS1JQWZVV0k0N0x1X1JVNkhHb2c&usp=drive_web#gid=0');
                        inSheet= sheet.getSheets()[0];
                        break;                  // var firstSemSheet= sheet.getSheetByName("firstSemester");//accessing insheets by name
                     
                                                                                                                
                        
                        case "Second semester result":
                        sheet= SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1o7qXTEI9ixK_Z-Afkb8WBGlj0r1yXEaham0rxdbzRow/edit#gid=736438352'); 
                        inSheet=sheet.getSheetByName("secondSemester");
                        break;
   
   
                        case "Third semester result":
                        sheet=SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1o7qXTEI9ixK_Z-Afkb8WBGlj0r1yXEaham0rxdbzRow/edit#gid=736438352');
                        inSheet=sheet.getSheetByName("thirdSemester");
                        break;
     
     
 }
       
  
  
  return(gettingCellValue(inSheet,rollno,columnName)); //returning the value returned by gettingcellvalue function
}



function gettingCellValue(sheet,rollno,columnName){  
 
  
 var indexFunction=gettingIndexValue(sheet,rollno,columnName);
  
 var indexObj=indexFunction();
  
  var rollRowIndex=indexObj.rollRowIndex;
  var columnNameIndex=indexObj.columnNameIndex;
  var values=getSheetProperty(sheet).values;  
    
  
  if(rollRowIndex ===null || columnNameIndex === null)
  {
    return "enter valid column name or roll number";  // if roll number or column name is wrong
  } 
   Logger.log(values[rollRowIndex][columnNameIndex]);
 
  
  if(!(values[rollRowIndex][columnNameIndex] === null))
  {
      return values[rollRowIndex][columnNameIndex];//the value at that particular id
  }

  
  else
  {
       return null;// if we have not enttered any cell value for that data in spreadsheet
  } 
    
  //Logger.log(values[rollRowIndex][columnNameIndex]);//just to cross verify my project remove my comments  
  
 //Logger.log(values[rollRowIndex][columnNameIndex]);
  
                                                      
                                                /* try to run with object method also later on.                    
                                               var myObj={
                                                                      headersRangeObj:headersRange;
                                                                       cellValuevalueObj:cellValue;
    
                                                                  }
 
                                                                         return myObj; */
}


                                                                 
function gettingHeaderValues(sem){  
  
  var sheet1;                                //opening the respective semester sheet on semester result basis  
  var inSheet1;
  
   switch(sem){
       
  case "First semester result":
                        sheet1= SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheet/ccc?key=0Aj2GJTiRHxZPdHRmbS1JQWZVV0k0N0x1X1JVNkhHb2c&usp=drive_web#gid=0');
                        inSheet1= sheet1.getSheets()[0];
                        break;                  // var firstSemSheet= sheet.getSheetByName("firstSemester");//accessing insheets by name
                     
                                                                                                                
                        
  case "Second semester result":
                        sheet1= SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1o7qXTEI9ixK_Z-Afkb8WBGlj0r1yXEaham0rxdbzRow/edit#gid=736438352'); 
                        inSheet1=sheet1.getSheetByName("secondSemester");
                        break;
   
   
 case "Third semester result":
                        sheet1=SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1o7qXTEI9ixK_Z-Afkb8WBGlj0r1yXEaham0rxdbzRow/edit#gid=736438352');
                        inSheet1=sheet1.getSheetByName("thirdSemester");
                        break;
   }
  
 return (getSheetProperty(inSheet1).headerRange); //calling getSheetProperty function which return object
}




/*function testingObject(){                //**for testing getSheetProperty function un comment this
  
  var sheet1= SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1o7qXTEI9ixK_Z-Afkb8WBGlj0r1yXEaham0rxdbzRow/edit#gid=736438352'); 
  var inSheet1=sheet1.getSheetByName("secondSemester");  //taking second semester sheet
  var testing=getSheetProperty(inSheet1).numColumn;
  Logger.log(testing);
 
}*/


function getSheetProperty(sheet)  //function is used to Access (total number of rows columns header values) by returning object
{

   var rnge = sheet.getDataRange();
   var numCol = rnge.getEndColumn() - rnge.getColumn() + 1; //ending column -startingcolumn +1 //returns an intiger
  var lastrow =sheet.getLastRow();
  var value=rnge.getValues();
  
  var headersRange = sheet.getRange(1, rnge.getColumn(), 1, numCol).getValues();
   
    
   
     var myObj={
       range:rnge,
       numColumn:numCol,
       lastRow:lastrow,  
       headerRange:headersRange,
       values:value
     }
     return myObj;
  
  

    
  }
     
// for testing gettingIndexValue function uncomment this
     
/*function testingClosure(){ 
   sheet1= SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1o7qXTEI9ixK_Z-Afkb8WBGlj0r1yXEaham0rxdbzRow/edit#gid=736438352'); 
                        inSheet1=sheet1.getSheetByName("secondSemester");
var funcn=gettingIndexValue(inSheet1,80012,"Enrollment Number");
var obj=funcn();
  Logger.log(obj.rollColumnIndex);
  Logger.log(obj.rollRowIndex);

}*/


function gettingIndexValue(sheet,rollNum,columnName){
  
  var columnNameIndex1=null; //very useful if function is not getting all arguments or we are not able to roll number, column number,and index values. 
  var rollColumnIndex1=null;                                           //****i should check this new concept of null in case any failure occurs 
  var rollRowIndex1=null;
  
  var headerArray=  getSheetProperty(sheet).headerRange;
  
  
  var indexValue=function(){
    
    
  
   if(columnName !== null)                        //getting column Index for the string "column name"  
   {
    
     for(var i=0;i<headerArray[0].length;i++)
    {
     if(columnName === headerArray[0][i])
     {
       columnNameIndex1 = i;
       break;
     }
   
    }
     
  }  
    
                                                          //getting column index for roll number
      for(var i=0;i<headerArray[0].length;i++)          
      {
        if("Roll Number" === headerArray[0][i])
        {
          rollColumnIndex1 = i;
         break; 
         
        } 
   
      }

  var rollvalue= sheet.getRange(1,((rollColumnIndex1)+1),getSheetProperty(sheet).lastRow).getValues();
    
  if(rollNum !== null)  
  {
    for(var i=0;i<getSheetProperty(sheet).lastRow;i++){  //getting location of roll number (mainly row value)
 
      if(rollNum === rollvalue[i][0])
     {
       rollRowIndex1=i;
       break;
     }
 
   }
 }   

    var myobj={
      
      columnNameIndex:columnNameIndex1,
      rollColumnIndex:rollColumnIndex1,
      rollRowIndex:rollRowIndex1
      
    }
    return myobj;
    
}
  return indexValue;

}

     
     
     
     
     
     
     
     
     
     //in progress
     /* var total = [];   //try to built sum method completely
   var sum;
    
   for(var i=0; i< lastRow; i++) //calculating total marks
   { 
          var ss=sheet.getRange(i+2,4, 1,4).getValues();
      // sum=0;
     
         for(var j=3;j<7;j++)    //****the subject marks should must be entered in 4th to 7th column necessary
     {
         Logger.log(ss[0][j]);
          
     }
    // total[i]=sum;
   }
    
    Logger.log(parseInt(total[1]));
  }  
  }
  }*/
   
