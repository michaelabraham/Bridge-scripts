#target bridge   
//written by Paul Riggott. Naming conventions edited and cancel button functionality added by Michael Abraham.

if( BridgeTalk.appName == "bridge" ) {  
if(!MenuElement.find("myMetaData")){
var newMenu = new MenuElement( "menu", "Witchcraft", "after Help", "myMetaData" );
}
var extMetadata = new MenuElement( "command", "Export Metadata", "at the end of myMetaData" , "xx1D" );
}
extMetadata.onSelect = function () { 
var win = new Window( 'dialog', 'Extract Data' ); 
win.alignChildren=["row","bottom"];
win.g1 = win.add('group');
win.title = win.g1.add('statictext',undefined,'Export Metadata');
win.title.alignment="bottom";
win.title.graphics.font = ScriptUI.newFont("Georgia","BOLDITALIC",20);
win.g0 = win.add('group');
win.g0.orientation= "row";
win.p1= win.g0.add("panel", undefined, undefined, {borderStyle:"black"}); 
win.p1.orientation= "row";
win.p1.preferredSize=[250,200];
win.grp10 = win.p1.add('group');
win.grp10.orientation= "column";
win.grp10.alignChildren = "left";
win.grp10.spacing=0;
win.grp10.cb1 = win.grp10.add('checkbox',undefined,'Keywords');
win.grp10.cb2 = win.grp10.add('checkbox',undefined,'Description');
win.grp10.cb3 = win.grp10.add('checkbox',undefined,'Headline');
win.grp10.cb4 = win.grp10.add('checkbox',undefined,'Title');
win.grp10.cb5 = win.grp10.add('checkbox',undefined,'Instructions');
win.grp10.cb6 = win.grp10.add('checkbox',undefined,'Date Created');
win.grp10.cb7 = win.grp10.add('checkbox',undefined,'State/Province');
win.grp10.cb8 = win.grp10.add('checkbox',undefined,'Job Identifier');
win.grp10.cb9 = win.grp10.add('checkbox',undefined,'Caption Writer');
win.grp10.cb10 = win.grp10.add('checkbox',undefined,'Width');
win.grp10.cb11 = win.grp10.add('checkbox',undefined,'Label');
win.grp20 = win.p1.add('group');
win.grp20.orientation= "column";
win.grp20.alignChildren = "left";
win.grp20.spacing=0;
win.grp20.cb7 = win.grp20.add('checkbox',undefined,'Location');
win.grp20.cb8 = win.grp20.add('checkbox',undefined,'City');
win.grp20.cb9 = win.grp20.add('checkbox',undefined,'Country');
win.grp20.cb10 = win.grp20.add('checkbox',undefined,'Rating');
win.grp20.cb11 = win.grp20.add('checkbox',undefined,'Copyrightinfo');
win.grp20.cb12 = win.grp20.add('checkbox',undefined,'Author');
win.grp20.cb13 = win.grp20.add('checkbox',undefined,'Camera Model');
win.grp20.cb14 = win.grp20.add('checkbox',undefined,'Camera Serial');
win.grp20.cb15 = win.grp20.add('checkbox',undefined,'Event');
win.grp20.cb16 = win.grp20.add('checkbox',undefined,'Height');
win.grp20.cb17 = win.grp20.add('checkbox',undefined,'File Size');
win.p2= win.add("panel", undefined, undefined, {borderStyle:"black"}); 
win.p2.preferredSize=[250,10];
win.grp30 = win.p2.add('group');
win.grp30.rb1 = win.grp30.add('radiobutton',undefined,'CSV File');
win.grp30.rb1.value=true;
win.grp30.rb2 = win.grp30.add('radiobutton',undefined,'TAB File');
win.g100 = win.add('group');
win.g100.bu1 = win.g100.add('button',undefined,'Get Metadata');
win.g100.bu1.preferredSize=[120,30];
win.g100.bu2 = win.g100.add('button',undefined,'Cancel');
win.g100.bu2.preferredSize=[120,30];
win.g100.bu2.onClick=function(){
    win.close(1);
    }
win.g100.bu1.onClick=function(){
    win.close(1);
    extractData();
    }
win.center();
var result = win.show();
function extractData(){
    var d = new Date();
        var month = d.getMonth();
        var day = d.getDate();
        var year = d.getFullYear();
        year = year.toString().substr(2,2);
        month = month + 1;
        month = month + "";
        if (month.length == 1){
            month = "0" + month;
            }
        day = day + '';
        if (day.length == 1){
            day = "0" + day;
            }
        var TodaysDate = year + month + day;
if(win.grp30.rb1.value){
var DataFile = File(Folder.desktop + "/" + TodaysDate.toString() + "_" + decodeURI(Folder(app.document.presentationPath).name) + ".csv");
var csvtab = ",";
}else{
    var DataFile = File(Folder.desktop + "/" + TodaysDate.toString() + "_" + decodeURI(Folder(app.document.presentationPath).name) + ".tab");
    var csvtab = "\t";
    }
var thumbs = app.document.selections; 
if(!thumbs.length) return;
DataFile.encoding = "UTF-8";
DataFile.open("w", "TEXT", "????");
$.os.search(/windows/i)  != -1 ? DataFile.lineFeed = 'windows'  : DataFile.lineFeed = 'macintosh';
var header = "FileName";
if(win.grp10.cb1.value)  header += csvtab + 'Keywords';
if(win.grp10.cb2.value)  header += csvtab +'Description';
if(win.grp10.cb3.value)  header += csvtab +'Headline';
if(win.grp10.cb4.value)  header += csvtab +'Title';
if(win.grp10.cb5.value)  header += csvtab +'Instructions';
if(win.grp10.cb6.value)  header += csvtab +'Date Created';
if(win.grp10.cb7.value)  header += csvtab +'State/Province';
if(win.grp10.cb8.value)  header += csvtab +'Job Identifier';
if(win.grp10.cb9.value)  header += csvtab +'Description Writer';
if(win.grp20.cb8.value)  header += csvtab +'Location';
if(win.grp20.cb8.value)  header += csvtab +'City';
if(win.grp20.cb9.value)  header += csvtab +'Country';
if(win.grp20.cb10.value)  header += csvtab +'Rating';
if(win.grp20.cb11.value)  header += csvtab +'Copyrightinfo';
if(win.grp20.cb12.value)  header += csvtab +'Author';
if(win.grp20.cb13.value)  header += csvtab +'Camera Model';
if(win.grp20.cb14.value)  header += csvtab +'Camera Serial';
if(win.grp20.cb15.value)  header += csvtab +'Event';
if(win.grp10.cb10.value) header += csvtab +'Width';
if(win.grp20.cb16.value) header += csvtab +'Height';
if(win.grp10.cb11.value) header += csvtab +'Label';
if(win.grp20.cb17.value) header += csvtab +'File Size';
DataFile.writeln(header);
for(var a =0;a<thumbs.length;a++){
if(thumbs[a].type != "file") continue;
var t = new Thumbnail(thumbs[a].spec);       
var md = t.synchronousMetadata;
var Line ='';
Line = decodeURI(thumbs[a].name);
if(win.grp10.cb1.value){
    var Keywords = md.read("http://ns.adobe.com/photoshop/1.0/","Keywords").toString().replace(/,/g,';');
    if(Keywords == undefined) Keywords ='';
    Line += csvtab + Keywords;
    }
if(win.grp10.cb2.value){
    var Description =  md.read("http://purl.org/dc/elements/1.1/","dc:description").toString().replace(/\n/g,'-');
    if(Description == undefined) Description ='';
     Line += csvtab + Description;
    }
if(win.grp10.cb3.value){
    var Headline = md.read("http://ns.adobe.com/photoshop/1.0/","photoshop:Headline");
    if(Headline == undefined) Headline ='';
    Line += csvtab + Headline;
    }
if(win.grp10.cb4.value){
    var Title = md.read("http://purl.org/dc/elements/1.1/","dc:title");
    if(Title == undefined) Title ='';
    Line += csvtab + Title;
    }
if(win.grp10.cb5.value){
    var Instructions = md.read("http://ns.adobe.com/photoshop/1.0/","photoshop:Instructions");
    if(Instructions == undefined) Instructions='';
    Line += csvtab + Instructions;
    }
if(win.grp10.cb6.value){
    try{
    var date = md.read("http://ns.adobe.com/photoshop/1.0/","photoshop:DateCreated").toString().substr(0, 10);
    }catch(e){date='';}
    Line += csvtab +  date;
    }
if(win.grp10.cb7.value){
        try{
    var state = md.read("http://ns.adobe.com/photoshop/1.0/","photoshop:State");
    }catch(e){state='';}
    Line += csvtab +  state;
}  
if(win.grp10.cb8.value){
        try{
    var jobID = md.read("http://ns.adobe.com/photoshop/1.0/","photoshop:TransmissionReference");
    }catch(e){jobID='';}
    Line += csvtab +  jobID;
} 
if(win.grp10.cb9.value){
        try{
    var cWriter = md.read("http://ns.adobe.com/photoshop/1.0/","photoshop:CaptionWriter");
    }catch(e){cWriter='';}
    Line += csvtab +  cWriter;
} 

if(win.grp20.cb7.value){
    var Location =  md.read("http://iptc.org/std/Iptc4xmpCore/1.0/xmlns/","Iptc4xmpCore:Location");
    if( Location == undefined)  Location ='';
    Line += csvtab +  Location;
    }
if(win.grp20.cb8.value){
    var City = md.read("http://ns.adobe.com/photoshop/1.0/","photoshop:City");
    if(City == undefined) City ='';
    Line += csvtab + City;
    }
if(win.grp20.cb9.value){
    var Country = md.read("http://ns.adobe.com/photoshop/1.0/","photoshop:Country");
    if(Country == undefined) Country ='';
    Line += csvtab + Country;
    }
if(win.grp20.cb10.value){
    var Rating = md.read("http://ns.adobe.com/xap/1.0/","Rating");
    if(Rating == undefined) Rating =0;
    Line += csvtab + Rating;
    }
if(win.grp20.cb11.value){
    var Copyrightinfo = md.read("http://purl.org/dc/elements/1.1/","dc:rights");
    if(Copyrightinfo == undefined) Copyrightinfo ='';
    Line += csvtab + Copyrightinfo;
    }
if(win.grp20.cb12.value){
    var Author = md.read("http://purl.org/dc/elements/1.1/","dc:creator");
    if(Author == undefined) Author = '';
    Line += csvtab + Author;
    }
if(win.grp20.cb13.value){
    var cameraModel = md.read("http://ns.adobe.com/tiff/1.0/","tiff:Model");
    if(cameraModel == undefined) cameraModel = '';
    Line += csvtab + cameraModel;
    }
if(win.grp20.cb14.value){
    var cameraSerial = md.read("http://ns.adobe.com/exif/1.0/aux/","aux:SerialNumber");
        if(cameraSerial == undefined) cameraSerial = '';
    Line += csvtab + cameraSerial;
    }
if(win.grp20.cb15.value){
    var event = md.read("http://iptc.org/std/Iptc4xmpExt/2008-02-29/","Iptc4xmpExt:Event");
        if(event == undefined) event = '';
    Line += csvtab + event;
    }
if(win.grp10.cb10.value){
    var Width = t.core.quickMetadata.width;
            if(Width == undefined) Width = '';
    Line += csvtab + Width;
    }
if(win.grp20.cb16.value) {
var Height = t.core.quickMetadata.height;
            if(Height == undefined) Height = '';
    Line += csvtab + Height;
    }
if(win.grp10.cb11.value){
    Line += csvtab + t.label;
    }
if(win.grp20.cb17.value) {
var Size = (t.spec.length/1024/1024).toFixed(2);
    Line += csvtab + Size;
    }
DataFile.writeln(Line);
    }
DataFile.close();
alert("File created on the desktop: " + decodeURI(DataFile.name));
}
};
